# This file is the actual code for the Python runnable set-sites-permissions
import dataiku
from dataiku.runnables import Runnable
from dss_constants import DSSConstants
from office365_client import Office365Session
from safe_logger import SafeLogger


logger = SafeLogger("sharepoint-tool plugin")


class SetSitesPermissionsRunnable(Runnable):
    """The base interface for a Python runnable"""
    # based on https://gist.github.com/ruanswanepoel/14fd1c97972cabf9ca3d6c0d9c5fc542

    def __init__(self, project_key, config, plugin_config):
        """
        :param project_key: the project in which the runnable executes
        :param config: the dict of the configuration of the object
        :param plugin_config: contains the plugin settings
        """
        logger.info('SharePoint Online plugin site select macro v{}'.format(DSSConstants.PLUGIN_VERSION))
        self.project_key = project_key
        self.config = config
        self.plugin_config = plugin_config
        connection_name = config.get("sharepoint_connection")
        client = dataiku.api_client()
        connection = client.get_connection(connection_name)
        connection_info = connection.get_info()
        self.tenant_id = connection_info.get("params", {}).get("tenantId")
        self.client_id = connection_info.get("params", {}).get("appId")
        credentials = connection_info.get_oauth2_credential()
        sharepoint_access_token = credentials.get("accessToken")
        self.sharepoint_url = config.get("sharepoint_url")
        self.roles = config.get("roles", {})
        self.session = Office365Session(access_token=sharepoint_access_token)
        
    def get_progress_target(self):
        """
        If the runnable will return some progress info, have this function return a tuple of 
        (target, unit) where unit is one of: SIZE, FILES, RECORDS, NONE
        """
        return None

    def run(self, progress_callback):
        """
        Do stuff here. Can return a string or raise an exception.
        The progress_callback is a function expecting 1 value: current progress
        """
        hostname, site_path = parse_sharepoint_site_url(self.sharepoint_url)
        site_id = self.get_sharepoint_site_id(hostname, site_path)
        json = {
            "roles": self.roles,
            "grantedToIdentities": [
                {
                    "application": {
                        "id": self.client_id,
                        "displayName": "Dataiku SharePoint plugin"
                    }
                }
            ]
        }
        url = "https://graph.microsoft.com/v1.0/sites/{}/permissions".format(
            site_id
        )
        response = self.session.request(method="POST", url=url, json=json)
        return "Your site id is {}<br/>Your tenant id is {}</br>Operation result: {}".format(
            site_id, self.tenant_id, response.json()
        )


    def get_sharepoint_site_id(self, hostname, site_path):
        url = "https://graph.microsoft.com/v1.0/sites/{}:/sites/{}?$select=id".format(
            hostname, site_path
        )
        response = self.session.get(url=url)
        json_response = response.json()
        return json_response.get("id")


def parse_sharepoint_site_url(sharepoint_url):
    # Todo:
    # - handle vanity URLs
    # - handle sub sites
    if not sharepoint_url:
        return None, None
    hostname = site_path = None
    url_tokens = sharepoint_url.split("/")
    if len(url_tokens)>3:
        if url_tokens[0] == "https:":
            hostname = url_tokens[2]
            site_path = url_tokens[4]
        else:
            hostname = url_tokens[0]
            site_path = url_tokens[2]
    else:
        return None, None
    return hostname, site_path
    