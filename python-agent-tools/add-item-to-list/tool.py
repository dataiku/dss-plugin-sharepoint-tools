import dataiku
from dataiku.llm.agent_tools import BaseAgentTool
from safe_logger import SafeLogger
from office365_client import Office365Session, Office365ListWriter
from dss_constants import DSSConstants

logger = SafeLogger("sharepoint-tool plugin")


class WriteToSharePointListTool(BaseAgentTool):

    def set_config(self, config, plugin_config):
        logger.info('SharePoint Online plugin list write tool v{}'.format(DSSConstants.PLUGIN_VERSION))

        self.config = config
        connection_name = config.get("sharepoint_connection")
        client = dataiku.api_client()
        connection = client.get_connection(connection_name)
        connection_info = connection.get_info()
        credentials = connection_info.get_oauth2_credential()
        sharepoint_access_token = credentials.get("accessToken")
        sharepoint_url = config.get("sharepoint_url")

        session = Office365Session(access_token=sharepoint_access_token)
        site_id, self.list_id = session.extract_site_list_from_url(sharepoint_url)
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
        site = session.get_site(site_id)
        self.list = site.get_list(self.list_id)
        self.output_schema, self.properties = self._get_schema_and_properties()

    def _get_schema_and_properties(self):
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
        output_schema = {
            "columns": output_columns
        }
        return output_schema, properties

    def get_descriptor(self, tool):
        logger.info("Properties detected on this list: {}".format(self.properties))
        return {
            "description": "This tool can be used to access lists on SharePoint Online. The input to this tool is a dictionary containing the new issue summary and description, e.g. '{'summary':'new issue summary', 'description':'new issue description'}'",
            "inputSchema": {
                "$id": "https://dataiku.com/agents/tools/search/input",
                "title": "Add an item to a SharePoint Online list tool",
                "type": "object",
                "properties": self.properties
            }
        }

    def invoke(self, input, trace):
        logger.info("Invoke with schema {}".format(self.output_schema))
        if self.output_schema is None:
            return {
                "error": "{}".format(self.initialization_error)
            }
        sharepoint_writer = Office365ListWriter(
            self.list,
            self.output_schema,
            write_from_dict=True
        )
        logger.info("Office365ListWriter initialised")
        row = input.get("input", {})

        # Log inputs and config to trace
        trace.span["name"] = "WRITE_TO_SHAREPOINT_LIST_TOOL_CALL"
        trace.inputs["list"] = self.list_id
        trace.inputs["row"] = row
        trace.attributes["config"] = {
            "sharepoint_connection_name": self.config.get("sharepoint_connection"),
            "sharepoint_connection_url": self.config.get("sharepoint_url")
        }

        logger.info("writing row")
        sharepoint_writer.write_row(row)
        logger.info("closing writer")
        sharepoint_writer.close()

        output_text = 'The record was added on the "{}" SharePoint list'.format(self.list_id)
        logger.info(output_text)

        # Log outputs to trace
        trace.outputs["output"] = output_text

        return {
            "output": output_text
        }
