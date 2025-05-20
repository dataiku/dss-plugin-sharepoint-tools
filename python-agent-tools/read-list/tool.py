import dataiku
from dataiku.llm.agent_tools import BaseAgentTool
from safe_logger import SafeLogger
from office365_client import Office365Session
from dss_constants import DSSConstants


logger = SafeLogger("sharepoint-tool plugin")


class SearchSharePointListTool(BaseAgentTool):

    def set_config(self, config, plugin_config):
        logger.info('SharePoint Online plugin search tool v{}'.format(DSSConstants.PLUGIN_VERSION))
        
        self.config = config
        connection_name = config.get("sharepoint_connection")
        client = dataiku.api_client()
        connection = client.get_connection(connection_name)
        connection_info = connection.get_info()
        credentials = connection_info.get_oauth2_credential()
        sharepoint_access_token = credentials.get("accessToken")
        sharepoint_url = config.get("sharepoint_url")

        self.session = Office365Session(access_token=sharepoint_access_token)
        site_id, self.list_id = self.session.extract_site_list_from_url(sharepoint_url)
        self.properties = {}
        self.output_schema = None
        self.initialization_error = None
        if not site_id:
            self.initialization_error = "The site in '{}' does not exists or is not accessible. Please check your credentials".format(
                sharepoint_url
            )
            return
        if not self.list_id:
            self.initialization_error = "The list in '{}' does not exists or is not accessible. Please check your credentials".format(
                sharepoint_url
            )
            return
        site = self.session.get_site(site_id)
        self.list = site.get_list(self.list_id)
        self.output_schema, self.properties = self._get_schema_and_properties()

    def _get_schema_and_properties(self):
        output_columns = []
        properties = {}
        for sharepoint_column in self.list.get_columns():
            column_description = sharepoint_column.get("description")
            if column_description:
                properties[sharepoint_column.get("name")] = {
                    "type": "string",  # we don't have access to that information
                    "description": column_description
                }
                output_columns.append({
                    "type": "string",  # we don't have access to that information
                    "name": sharepoint_column.get("name")
                })
        output_schema = {
            "columns": output_columns
        }
        return output_schema, properties

    def get_descriptor(self, tool):
        # we want to modify the description to add the columns decriptions retrieved from sharepoint
        return {
            "description": "This tool can be used to access lists on SharePoint Online. The input to this tool is a dictionary containing the name of the column to search and the term to search in it, e.g. '{'City':'Paris', 'Urgency':'High'}'",
            "inputSchema": {
                "$id": "https://dataiku.com/agents/tools/search/input",
                "title": "Search a SharePoint Online list tool",
                "type": "object",
                "properties": self.properties
            }
        }

    def invoke(self, input, trace):
        if self.initialization_error:
            return {"error": self.initialization_error}

        args = input.get("input", {})
        filter_tokens = []
        for arg in args:
            filter = "fields/{} eq '{}'".format(arg, args.get(arg))
            filter_tokens.append(filter)
        filter = " and ".join(filter_tokens)
        hits = []
        try:
            for row in self.list.get_next_row(filter=filter):
                fields = row.get("fields", {})
                filtered_fields = self._filter_fields(fields)
                hits.append(filtered_fields)
        except Exception as error:
            logger.error("Error {}".format(error))
            return {"error": "There was an error while searching SharePoint Online"}

        return {
            "output": hits
        }

    def _filter_fields(self, fields):
        filtered_fields = {}
        for field in fields:
            if field in self.properties:
                filtered_fields[field] = fields.get(field)
        return filtered_fields
