# SharePoint Tools plugin

This plugin provide tools for allowing LLM agents to interact with SharePoint Online. It requires Dataiku >= 12.6.2

## Setup

An [active SharePoint connection](https://doc.dataiku.com/dss/latest/connecting/sharepoint-online.html) has to be available.

### Example Use Case: Create List Item Based on User Input

In this example, we aim to create a simple ticketing service. We will process a dataset of messages from users and create a SharePoint Online list item describing the issue, if necessary.
- First, create the target SharePoint list for the service. Create all the columns that need to be populated by Dataiku. Make the name of each column clear. Then, for each column, add a description (click on the **down arrow** on the right of the column name > **Column settings** > **Edit** > **Description** ). It should explain exactly what information goes into the column. This set of descriptions will be used by the LLM to what information goes in wich column.
- On Dataiku, create the agent tool. In the project's flow, click on **Analysis** > **Agent tools** > **+New agent tool** > **Create a SharePoint Online list item**. Name the identifier and the agent.
- Then, create the actual agent. Click on **+Other** > **Generative AI** > **Visual Agent**, name the agent, click on v1, and add the tool (**+Add tool**) created in the previous step. In the agent's *Additional prompt* section, describe your specifications for what the parameters in the ticket should contain. The agent tool will fill in the parameters named after each column present in the target list.
- Once this is done, the agent can be used in place of an LLM in your prompt and LLM settings. For instance, a simple LLM recipe with the prompt *You are an IT agent and your mission is to write a SharePoint list item if the user's message requires it.*

### License

Copyright 2025 Dataiku SAS

This plugin is distributed under the Apache License version 2.0