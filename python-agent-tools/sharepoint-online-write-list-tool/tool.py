import dataiku
from dataiku.llm.agent_tools import BaseAgentTool
from safe_logger import SafeLogger
from office365_client import Office365Session, Office365ListWriter

logger = SafeLogger("sharepoint-tool plugin")


class WriteToSharePointListTool(BaseAgentTool):

    def set_config(self, config, plugin_config):
        logger.info('SharePoint Online plugin list write tool v{}'.format("0.0.1"))

        connection_name = config.get("sharepoint_connection")
        client = dataiku.api_client()
        connection = client.get_connection(connection_name)
        connection_info = connection.get_info()
        credentials = connection_info.get_oauth2_credential()
        sharepoint_access_token = credentials.get("accessToken")
        sharepoint_url = config.get("sharepoint_url")

        session = Office365Session(access_token=sharepoint_access_token)
        site_id, list_id = session.extract_site_list_from_url(sharepoint_url)
        site = session.get_site(site_id)
        self.list = site.get_list(list_id)
        self.output_schema = None

    def get_descriptor(self, tool):
        output_columns = []
        properties = {}
        required = []
        for sharepoint_column in self.list.get_columns():
            column_description = sharepoint_column.get("description")
            column_description = sharepoint_column.get("description")
            if column_description:
                properties[sharepoint_column.get("name")] = {
                    "type": "string",  # we don't have access to that information
                    "name": sharepoint_column.get("name")
                }
                required.append(sharepoint_column.get("name"))  # For now...
                output_columns.append({
                    "type": "string",  # we don't have access to that information
                    "name": sharepoint_column.get("name")
                })
        self.output_schema = {
            "columns": output_columns
        }
        logger.info("Properties detected on this list: {}".format(properties))
        return {
            "description": "This tool can be used to access lists on SharePoint Online. The input to this tool is a dictionary containing the new issue summary and description, e.g. '{'summary':'new issue summary', 'description':'new issue description'}'",
            "inputSchema": {
                "$id": "https://dataiku.com/agents/tools/search/input",
                "title": "Add an item to a SharePoint Online list tool",
                "type": "object",
                "properties": properties
            }
        }

    def invoke(self, input, trace):
        sharepoint_writer = Office365ListWriter(
            self.list,
            self.output_schema,
            write_from_dict=True
        )
        row = input.get("input", {})
        sharepoint_writer.write_row(row)
        sharepoint_writer.close()

        return {
            "output": 'The record was added on the "{}" SharePoint list'.format(self.list)
        }
