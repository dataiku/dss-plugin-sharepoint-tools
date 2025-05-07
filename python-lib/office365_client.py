import requests
import time
from safe_logger import SafeLogger
from office365_site import Office365Site
from office365_drive import Office365Drive
from office365_messages import Office365Messages
from office365_auth import Office365Auth
from office365_commons import get_next_page_url, get_error, prepare_row, is_throttling, get_retry_after_value
from dss_constants import DSSConstants


logger = SafeLogger("office-365 plugin", [])


class Office365Session():
    def __init__(self, access_token=None):
        self.session = requests.Session()
        self.session.auth = Office365Auth(access_token=access_token)
        self.is_batch_mode = False
        self.requests_buffer = []
        self.batch_size = 0

    def request(self, **kwargs):
        raise_on = kwargs.pop("raise_on", {})
        cannot_raise = kwargs.pop("cannot_raise", False)
        force_no_batch = kwargs.pop("force_no_batch", False)

        if self.is_batch_mode and not force_no_batch:
            self.requests_buffer.append(kwargs)
            if len(self.requests_buffer) >= self.batch_size:
                self.flush()
            return

        should_retry = True
        while should_retry:
            should_retry = False
            response = self.session.request(**kwargs)
            if is_throttling(response):
                retry_after = get_retry_after_value(response)
                if retry_after:
                    logger.warning("SharePoint is throttling... Sleeping for {} seconds".format(retry_after))
                    time.sleep(retry_after)
                    should_retry = True
                    logger.warning("Retrying")
        error_message = get_error(response)
        if raise_on:
            status_code = response.status_code
            error_message = raise_on.get(status_code)
            if error_message:
                raise Exception(error_message)
        if error_message and not cannot_raise:
            raise Exception(error_message)
        return response

    def get(self, **kwargs):
        kwargs["method"] = "GET"
        response = self.request(**kwargs)
        error_message = get_error(response)
        if error_message and not kwargs.get("cannot_raise"):
            raise Exception(error_message)
        return response

    def get_headers(self):
        headers = DSSConstants.JSON_HEADERS
        return headers

    def get_item(self, **kwargs):
        kwargs["headers"] = kwargs.get("headers", {})
        kwargs["headers"].update(DSSConstants.JSON_HEADERS)
        kwargs["headers"].update(DSSConstants.GZIP_HEADERS)
        kwargs["cannot_raise"] = True
        response = self.get(
            **kwargs
        )
        status_code = response.status_code
        if status_code == 404:
            return {}
        json_response = response.json()
        return json_response

    def get_next_item(self, **kwargs):
        kwargs["headers"] = kwargs.get("headers", {})
        kwargs["headers"].update(DSSConstants.JSON_HEADERS)
        kwargs["headers"].update(DSSConstants.GZIP_HEADERS)
        is_first_get = True
        next_page_url = None
        while next_page_url or is_first_get:
            kwargs["url"] = next_page_url or kwargs["url"]
            if next_page_url:
                # As next_page_url already contains query params
                kwargs["params"] = None
            response = self.get(
                **kwargs
            )
            is_first_get = False
            json_response = response.json()
            next_page_url = get_next_page_url(json_response)
            items = json_response.get("value", [])
            for item in items:
                yield item

    def get_next_site(self):
        for site in self.get_next_item(
                url=self.get_sites_url(),
                params={"search": "*"}
        ):
            yield site

    def get_my_tasks(self):
        # https://graph.microsoft.com/v1.0/me/planner/tasks
        for task in self.get_next_item(url=self.get_endpoint_url_for("me/planner/tasks")):
            yield task

    def get_next_task(self, plan_id):
        for task in self.get_next_item(url=self.get_endpoint_url_for("planner/plans/{}/tasks".format(plan_id))):
            yield task

    def get_next_plan(self, group_id):
        for plan in self.get_next_item(
            url=self.get_endpoint_url_for("/groups/{}/planner/plans".format(group_id))
        ):
            yield plan

    def get_all_items(self, **kwargs):
        items = []
        for item in self.get_next_item(**kwargs):
            items.append(item)
        return items

    def start_batch_mode(self, batch_size=None):
        batch_size = batch_size or DSSConstants.DEFAULT_BATCH_SIZE
        self.is_batch_mode = True
        self.batch_size = batch_size
        self.requests_buffer = []

    def close(self):
        self.flush()
        self.is_batch_mode = False

    def flush(self):
        responses = self.process_batch(self.requests_buffer)
        self.requests_buffer = []
        assert_responses_ok(responses)

    def get_site(self, site_id):
        return Office365Site(self, site_id)

    def get_messages(self, search_space=None):
        return Office365Messages(self, search_space=search_space)

    def get_site_id(self, site_name):
        search_by_web_url = True if "/" in site_name else False
        for site in self.get_next_site():
            if search_by_web_url:
                full_site_path = "/".join(site.get("webUrl").split("/")[3:])
                if full_site_path == site_name:
                    return site.get("id")
            else:
                if site.get("name") == site_name:
                    return site.get("id")
        return None

    def get_drive(self, drive_id):
        return Office365Drive(self, drive_id)

    def process_batch(self, requests_buffer):
        if not requests_buffer:
            return {}
        data = {}
        requests = []
        counter = 1
        for request_kwargs in requests_buffer:
            request = {
                "id": "{}".format(counter),
                "method": request_kwargs.get("method"),
                "url": self.get_relative_url(request_kwargs.get("url")),
            }
            if request_kwargs.get("headers"):
                request["headers"] = request_kwargs.get("headers")
            if request_kwargs.get("json"):
                request["body"] = request_kwargs.get("json")
            if request_kwargs.get("data"):
                request["data"] = request_kwargs.get("data")
            requests.append(
                request
            )
            counter += 1
        data["requests"] = requests
        response = self.session.request(
            method="POST",
            url=self.get_batch_url(),
            headers=DSSConstants.JSON_HEADERS,
            json=data
        )
        status_code = response.status_code
        if status_code >= 400:
            error_message = "Batch error {}".format(status_code)
            try:
                json_response = response.json()
                error_message += ". {}".format(json_response.get("error").get("message"))
            except Exception as sub_error_message:
                logger.debug("Could not enrich error message {}".format(sub_error_message))
            logger.error("Error {}, dumping content: {}".format(status_code, response.content))
            raise Exception("Error {}".format(status_code))
        json_response = response.json()
        return json_response.get("responses", {})

    def get_batch_url(self):
        return "/".join(
            [
                self.get_endpoint_url(),
                "$batch"
            ]
        )

    def get_sites_url(self):
        return "/".join(
            [
                self.get_endpoint_url(),
                "sites"
            ]
        )

    def get_relative_url(self, full_url):
        url_base = self.get_endpoint_url()
        relative_url = full_url
        if full_url.startswith(url_base):
            relative_url = full_url.replace(url_base, "")
        return relative_url

    def get_endpoint_url(self):
        return "https://graph.microsoft.com/v1.0"

    def get_endpoint_url_for(self, root_path):
        return "/".join(
            [
                self.get_endpoint_url(),
                root_path
            ]
        )

    def extract_site_list_from_url(self, url):
        logger.info("searching site from url {}".format(url))
        url_tokens = url.strip("/").split("/")
        list_name = url_tokens[-2:-1][0]
        site_name = url_tokens[-4:-3][0]
        site_id = None
        logger.info("searching for site '{}'".format(site_name))
        site_counter = 0
        for site in self.get_next_site():
            if site.get("name") == site_name:
                site_id = site.get("id")
            site_counter += 1
        logger.info("searched through {} sites, id found is '{}'".format(site_counter, site_id))
        if not site_id:
            raise Exception("The site '{}' was not found. Check the URL or your credentials.".format(site_name))
        site = self.get_site(site_id)
        list_id = None
        logger.info("searching for list '{}'".format(list_name))
        list_counter = 0
        for sharepoint_list in site.get_next_list():
            if sharepoint_list.get("name") == list_name:
                list_id = sharepoint_list.get("id")
            list_counter += 1
        logger.info("searched through {} lists, id found is '{}'".format(list_counter, list_id))
        if not list_id:
            raise Exception("The list '{}' was not found. Check the URL or your credentials.".format(list_name))
        return site_id, list_id

    def search_list(self, query):
        """
        POST /search/query
        Content-Type: application/json

        {
        "requests": [
            {
            "entityTypes": [
                "driveItem"  # should be listItem ?
            ],
            "query": {
                "queryString": "contoso"
            }
            }
        ]
        }
        """
        json = {
            "requests": [
                {
                    "entityTypes": [
                        "listItem"
                    ],
                    "query": {
                        "queryString": "{}".format(query)
                    }
                }
            ]
        }
        url = "/".join(
            [
                self.get_endpoint_url(),
                "search",
                "query"
            ]
        )
        response = self.request(method="POST", url=url, json=json)
        json_response = response.json()
        value = json_response.get("value", [])
        if len(value) > 0:
            hits_containers = value[0].get("hitsContainers", [])
            if len(hits_containers) > 0:
                hits = hits_containers.get("hits", [])
                return hits
        return []


def get_relative_url(url_base, full_url):
    relative_url = full_url
    if full_url.startswith(url_base):
        relative_url = full_url.replace(url_base, "")
    return relative_url


def assert_responses_ok(responses):
    max_retry_after = 0
    for response in responses:
        if int(response.get("status", 200)) >= 400:
            logger.error("Error during batch, dumping responses: {}".format(responses))
            raise Exception("Batch id {} failed with error {}. {} / {}".format(
                response.get("id"),
                response.get("status"),
                response.get("body"),
                response.get("header")
            ))
        retry_after = response.get("header", {}).get("Retry-After", 0)
        if retry_after > max_retry_after:
            max_retry_after = retry_after
    return True


class Office365ListWriter(object):
    def __init__(self, list, dataset_schema, batch_size=None, write_from_dict=False):
        self.list = list
        self.list.session.start_batch_mode(batch_size=batch_size)
        self.columns = dataset_schema.get("columns")
        self.write_from_dict = write_from_dict

    def write_row(self, row):
        if self.write_from_dict:
            self.list.write_row(row)
        else:
            self.list.write_row(prepare_row(row, self.columns))

    def close(self):
        self.list.session.close()
