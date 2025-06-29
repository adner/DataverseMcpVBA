# Using Excel as an MCP client

This repo contains a VBA module that can be used in Excel to connect to the Dataverse MCP Server running locally.

See instructions [here](https://learn.microsoft.com/en-us/power-apps/maker/data-platform/data-platform-mcp) for how to set up the MCP Server.

The VBA code relies on a JSON converter that can be found [here](https://github.com/VBA-tools/VBA-JSON/blob/master/JsonConverter.bas).

The ConnectionUrl and TenantId in the code needs to be manually updated.

The following functions are available and can be invoked  from a cell in Excel:

- **DvMcpListTools** - Lists all tools available from the Dataverse MCP Server.
  - Example: `=DvMcpListTools()`
- **DvMcpUpdateRecord** - Updates a record. Parameters:
  - tablename
  - GUID of record
  - JSON-object
  -  Example: `=DvMcpUpdateRecord("contact";"[GUID]";"{'firstname':'" & E2 &"'}")`
- **DvMcpCreateRecord** - Creates a record. Parameters:
  - tablename
  - JSON-object
  - Example: `=DvMcpCreateRecord("contact";"{'firstname':'Testperson', 'lastname':'" & G30 &"'}")`
