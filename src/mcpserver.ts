import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { callDynamicsApi } from './dynamicsApiClient.js';
import { DynamicsEntityManager } from './dynamicsEntityManager.js';
import { RequestHandlerExtra } from '@modelcontextprotocol/sdk/shared/protocol.js';
import { ServerRequest, ServerNotification } from '@modelcontextprotocol/sdk/types.js';

const entityManager = new DynamicsEntityManager();
const DEFAULT_PAGE_SIZE = 5;

/**
 * Safely send notifications
 */
async function safeNotification(
    context: RequestHandlerExtra<ServerRequest, ServerNotification>,
    notification: ServerNotification
): Promise<void> {
    try {
        await context.sendNotification(notification);
    } catch (error) {
        console.log('Notification failed (this is normal in test environments):', error);
    }
}

/**
 * Build OData $filter string from an object
 */
function buildFilterString(filterObject?: Record<string, string>): string | null {
    if (!filterObject || Object.keys(filterObject).length === 0) {
        return null;
    }
    return Object.entries(filterObject)
        .map(([key, value]) => `${key} eq '${value}'`)
        .join(' and ');
}

// --- Zod Schemas for Tool Arguments ---
const odataQuerySchema = z.object({
    entity: z.string().describe("The OData entity set to query (e.g., CustomersV3, ReleasedProductsV2)."),
    select: z.string().optional().describe("OData $select query parameter to limit the fields returned."),
    filter: z.record(z.string()).optional().describe("Key-value pairs for filtering."),
    expand: z.string().optional().describe("OData $expand query parameter."),
    top: z.number().optional().describe(`The number of records to return per page. Defaults to ${DEFAULT_PAGE_SIZE}.`),
    skip: z.number().optional().describe("The number of records to skip for pagination."),
    crossCompany: z.boolean().optional().describe("Set to true to query across all companies."),
});

const createCustomerSchema = z.object({
    customerData: z.record(z.unknown()).describe("A JSON object for the new customer."),
});

const updateCustomerSchema = z.object({
    dataAreaId: z.string().describe("The dataAreaId of the customer (e.g., 'usmf')."),
    customerAccount: z.string().describe("The customer account ID to update (e.g., 'PM-001')."),
    updateData: z.record(z.unknown()).describe("A JSON object with the fields to update."),
});

const getEntityCountSchema = z.object({
    entity: z.string().describe("The OData entity set to count (e.g., CustomersV3)."),
    crossCompany: z.boolean().optional().describe("Set to true to count across all companies."),
});

const createSystemUserSchema = z.object({
     userData: z.record(z.unknown()).describe("A JSON object for the new system user. Must include UserID, Alias, Company, etc."),
});

const assignUserRoleSchema = z.object({
    associationData: z.record(z.unknown()).describe("JSON object for the role association. Must include UserId and SecurityRoleIdentifier."),
});

const exportToPackageSchema = z.object({
    definitionGroupId: z.string().describe("Definition group name for execution"),
    packageName: z.string().describe("Package Name."),
    executionId: z.string().describe("Execution Id for export.")
    .default(() => {
        const now = new Date();
        return `EXEC-${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}-${now.getHours().toString().padStart(2, '0')}${now.getMinutes().toString().padStart(2, '0')}${now.getSeconds().toString().padStart(2, '0')}`;
    }),
    legalEntityId: z.string().describe("Company for export."),
});

const executionSchema = z.object({
    executionId: z.string().describe("Execution Id of the data management job."),
});

const uniqueFileNameSchema = z.object({
    uniqueFileName: z.string().describe("Unique file name of the azure writable URL."),
});

const importFromPackageSchema = z.object({
    packageUrl: z.string().describe("The URL of the package to import."),
    definitionGroupId: z.string().describe("Definition group name for execution"),
    executionId: z.string().describe("Execution Id for export.")
    .default(() => {
        const now = new Date();
        return `EXEC-${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}-${now.getHours().toString().padStart(2, '0')}${now.getMinutes().toString().padStart(2, '0')}${now.getSeconds().toString().padStart(2, '0')}`;
    }),
    legalEntityId: z.string().describe("Company for export."),
});


/**
 * Creates and configures the MCP server with all the tools for the D365 API.
 * @returns {McpServer} The configured McpServer instance. 
 */
export const getServer = (): McpServer => {
    const server = new McpServer({
        name: 'd365-fno-mcp-server',
        version: '1.0.0',
    });

    // --- Tool Definitions ---

    server.tool(
        'odata-Query',
        'Executes a generic GET request against a Dynamics 365 OData entity. The entity name does not need to be case-perfect. Responses are paginated.',
        odataQuerySchema.shape,
        async (args: z.infer<typeof odataQuerySchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            
            const correctedEntity = await entityManager.findBestMatchEntity(args.entity);
    
            if (!correctedEntity) {
                return {
                    isError: true,
                    content: [{ type: 'text', text: `Could not find a matching entity for '${args.entity}'. Please provide a more specific name.` }]
                };
            }
            
            const effectiveArgs = { ...args };

            if (effectiveArgs.filter?.dataAreaId && effectiveArgs.crossCompany !== false) {
                if (!effectiveArgs.crossCompany) {
                    await safeNotification(context, {
                        method: "notifications/message",
                        params: { level: "info", data: `Filter on company ('dataAreaId') detected. Automatically enabling cross-company search.` }
                    });
                }
                effectiveArgs.crossCompany = true;
            }

            await safeNotification(context, {
                method: "notifications/message",
                params: { level: "info", data: `Corrected entity name from '${args.entity}' to '${correctedEntity}'.` }
            });
            
            const { entity, ...queryParams } = effectiveArgs;
            const filterString = buildFilterString(queryParams.filter);
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/${correctedEntity}`);

            // PAGINATION: Apply query parameters including the new skip and a default top.
            const topValue = queryParams.top || DEFAULT_PAGE_SIZE;
            url.searchParams.append('$top', topValue.toString());

            if (queryParams.skip) {
                url.searchParams.append('$skip', queryParams.skip.toString());
            }

            if (queryParams.crossCompany) url.searchParams.append('cross-company', 'true');
            if (queryParams.select) url.searchParams.append('$select', queryParams.select);
            if (filterString) url.searchParams.append('$filter', filterString);
            if (queryParams.expand) url.searchParams.append('$expand', queryParams.expand);
            
            return callDynamicsApi('GET', url.toString(), null, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'create-Customer',
        'Creates a new customer record in CustomersV3.',
        createCustomerSchema.shape,
        async ({ customerData }: z.infer<typeof createCustomerSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/CustomersV3`;
            return callDynamicsApi('POST', url, customerData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'update-Customer',
        'Updates an existing customer record in CustomersV3 using a PATCH request.',
        updateCustomerSchema.shape,
        async ({ dataAreaId, customerAccount, updateData }: z.infer<typeof updateCustomerSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/CustomersV3(dataAreaId='${dataAreaId}',CustomerAccount='${customerAccount}')`;
            return callDynamicsApi('PATCH', url, updateData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'get-Entity-Count',
        'Gets the total count of records for a given OData entity.',
        getEntityCountSchema.shape,
        async ({ entity, crossCompany }: z.infer<typeof getEntityCountSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
             const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/${entity}/$count`);
             if (crossCompany) url.searchParams.append('cross-company', 'true');
             return callDynamicsApi('GET', url.toString(), null, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'create-SystemUser',
        'Creates a new user in SystemUsers.',
        createSystemUserSchema.shape,
        async ({ userData }: z.infer<typeof createSystemUserSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/SystemUsers`;
            return callDynamicsApi('POST', url, userData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'assign-User-Role',
        'Assigns a security role to a user in SecurityUserRoleAssociations.',
        assignUserRoleSchema.shape,
        async ({ associationData }: z.infer<typeof assignUserRoleSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/SecurityUserRoleAssociations`;
            return callDynamicsApi('POST', url, associationData as Record<string, unknown>, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'action-initialize-DataManagement',
        'Executes the InitializeDataManagement action on the DataManagementDefinitionGroups entity.',
        z.object({}).shape,
        async (_args: {}, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.InitializeDataManagement`;
            return callDynamicsApi('POST', url, {}, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        'get-OData-Metadata',
        'Retrieves the OData $metadata document for the service.',
        z.object({}).shape,
        async (_args: {}, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
             const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/$metadata`;
             return callDynamicsApi('GET', url.toString(), null, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        "export-to-package", 
        "Exports data to a package, using the provided definition group, package name, and execution ID.",
        exportToPackageSchema.shape,
        async ({ definitionGroupId, packageName, executionId, legalEntityId }: z.infer<typeof exportToPackageSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>
        ) => {
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.ExportToPackage`);
            return callDynamicsApi('POST', url.toString(), {
                definitionGroupId,
                packageName,
                executionId,
                reExecute: true,
                legalEntityId
            }, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        "get-datamanagement-executionstatus", 
        "Get data managmenet export package or job status, using the provided execution ID.",
        executionSchema.shape,
        async ({ executionId }: z.infer<typeof executionSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.GetExecutionSummaryStatus`);
            return callDynamicsApi('POST', url.toString(), { executionId }, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        "get-datamanagement-exportedpackageURL", 
        "Retrieves the URL of the exported package, using the provided execution ID.",
        executionSchema.shape,
        async ({ executionId }: z.infer<typeof executionSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.GetExportedPackageUrl`);
            return callDynamicsApi('POST', url.toString(), { executionId }, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        "get-datamanagement-executionErrors", 
        "Retrieves the execution errors, using the provided execution ID.",
        executionSchema.shape,
        async ({ executionId }: z.infer<typeof executionSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.GetExecutionErrors`);
            return callDynamicsApi('POST', url.toString(), { executionId }, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        "get-azure-writable-url", 
        "Get the Azure writable URL for a data management import package, using the unique file name.",
        uniqueFileNameSchema.shape,
        async ({ uniqueFileName }: z.infer<typeof uniqueFileNameSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.GetAzureWriteUrl`);
            return callDynamicsApi('POST', url.toString(), { uniqueFileName }, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );

    server.tool(
        "import-from-package-async", 
        "Imports data from a package asynchronously, using the provided definition group, package name, and execution ID.",
        importFromPackageSchema.shape,
        async ({ packageUrl, definitionGroupId, executionId, legalEntityId }: z.infer<typeof importFromPackageSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>
        ) => {
            const url = new URL(`${process.env.DYNAMICS_RESOURCE_URL}/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities.ImportFromPackageAsync`);
            return callDynamicsApi('POST', url.toString(), {
                packageUrl,
                definitionGroupId,
                executionId,
                execute: true,
                overWrite: true,
                legalEntityId,
                failOnError: true,
                runAsyncWithoutBatch: true,
                thresholdToRunInBatch:1
            }, async (notification) => {
                await safeNotification(context, notification);
            });
        }
    );
    

    //Added by JP Start
const createSalesOrderHeaderSchema = z.object({
  dataAreaId: z.string().describe("Legal entity, e.g. 'usmf'."),
  RequestedShippingDate: z.string().describe("ISO 8601 date, e.g. '2025-10-20'."),
  OrderingCustomerAccountNumber: z.string().describe("Ordering customer account number, e.g. 'US-001'."),
  SalesOrderNumber: z.string().optional().describe("Optional. If omitted, D365 auto-assigns.")
});

const addSalesOrderLineSchema = z.object({
  dataAreaId: z.string().describe("Legal entity, e.g. 'usmf'."),
  SalesOrderNumber: z.string().describe("Existing sales order number to add the line to."),
  ItemNumber: z.string().describe("Item number."),
  // Accept 5 or "5"
  OrderedSalesQuantity: z.coerce.number().describe("Ordered quantity."),
  // Optional; system may default if omitted
  SiteId: z.string().optional().describe("Optional site.")
});
//Added by JP End

//Added by JP Start
// --- Create Sales Order Header ---
server.tool(
  'createSalesOrderHeader',
  'Creates a sales order header in SalesOrderHeadersV4. Omit SalesOrderNumber to auto-assign.',
  createSalesOrderHeaderSchema.shape,
  async (args: z.infer<typeof createSalesOrderHeaderSchema>, context) => {
    const payload: Record<string, unknown> = {
      dataAreaId: args.dataAreaId,
      RequestedShippingDate: args.RequestedShippingDate,
      OrderingCustomerAccountNumber: args.OrderingCustomerAccountNumber,
      ...(args.SalesOrderNumber ? { SalesOrderNumber: args.SalesOrderNumber } : {})
    }; 

    await safeNotification(context, {
      method: "notifications/message",
      params: { level: "info", data: `Creating sales order header${args.SalesOrderNumber ? `: ${args.SalesOrderNumber}` : ''}` }
    });

    const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/SalesOrderHeadersV4`;
    const result = await callDynamicsApi('POST', url, payload, async (notification) => { await safeNotification(context, notification); });

    // Try to surface the assigned number (if auto-numbered)
    let assigned: string | undefined;
    const txt = result.content?.[0]?.type === 'text' ? result.content[0].text : '';
    try {
      const obj = JSON.parse(txt);
      assigned = obj?.SalesOrderNumber ?? obj?.SalesId ?? args.SalesOrderNumber;
    } catch { /* ignore non-JSON */ }

    return {
      ...(result.isError ? { isError: true } : {}),
      content: [{
        type: 'text',
        text: assigned
          ? `Header created. SalesOrderNumber: ${assigned}\n\nRaw response:\n${txt}`
          : (txt || 'Header created.')
      }]
    };
  }
);

// --- Add a Single Sales Order Line ---
server.tool(
  'addSalesOrderLine',
  'Adds a single line to an existing order in SalesOrderLinesV3.',
  addSalesOrderLineSchema.shape,
  async (args: z.infer<typeof addSalesOrderLineSchema>, context) => {
    await safeNotification(context, {
      method: "notifications/message",
      params: { level: "info", data: `Adding line to ${args.SalesOrderNumber}: ${args.ItemNumber} x ${args.OrderedSalesQuantity}` }
    });

    const payload: Record<string, unknown> = {
      dataAreaId: args.dataAreaId,
      SalesOrderNumber: args.SalesOrderNumber,
      ItemNumber: args.ItemNumber,
      OrderedSalesQuantity: args.OrderedSalesQuantity,
      ...(args.SiteId ? { SiteId: args.SiteId } : {})
    };

    const url = `${process.env.DYNAMICS_RESOURCE_URL}/data/SalesOrderLinesV3`;
    const res = await callDynamicsApi('POST', url, payload, async (notification) => { await safeNotification(context, notification); });

    const first = res.content?.[0];
    const text = (first && first.type === 'text') ? first.text : JSON.stringify(res);

    return {
      ...(res.isError ? { isError: true } : {}),
      content: [{ type: 'text', text }]
    };
  }
);


//Added by JP End

 

 // --- Schema for Dequeue Tool ---
const dequeueJobSchema = z.object({
    activityId: z.string().describe("The activity ID of the D365FO job to dequeue."),
    entity: z.string().describe("The entity name associated with the job (e.g., 'Customer groups').")
});

// --- Tool Definition ---
server.tool(
    "dequeue-job",
    "Dequeues a D365FO job using an activity ID and entity name.",
    dequeueJobSchema.shape,
    async ({ activityId, entity }: z.infer<typeof dequeueJobSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
        try {
            // Encode entity name (spaces → %20)
            const encodedEntity = encodeURIComponent(entity);

            // Build URL
            const url = new URL(
                `${process.env.DYNAMICS_RESOURCE_URL}/api/connector/dequeue/${activityId}?entity=${encodedEntity}`
            );

            // Call D365 API
            return await callDynamicsApi('GET', url.toString(), null, async (notification) => {
                await safeNotification(context, notification);
            });
        } catch (error) {
            return {
                isError: true,
                content: [{ type: "text", text: `Failed to dequeue job: ${String(error)}` }]
            };
        }
    }
);
//=================================

// --- Schema for Download Tool ---
const downloadJobSchema = z.object({
    downloadLocation: z.string().describe(
        "The full download URL including query params. Example: https://orgd8ac6da4.operations.dynamics.com/api/connector/download/{activityId}?correlation-id={cid}&blob={blob}"
    )
});

// --- Tool Definition ---
server.tool(
    "download-RI-job-package",
    "Downloads a D365FO job package from the provided download location from RI deque job.",
    downloadJobSchema.shape,
    async ({ downloadLocation }: z.infer<typeof downloadJobSchema>, context: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
        try {
            // Use downloadLocation directly
            return await callDynamicsApi('GET', downloadLocation, null, async (notification) => {
                await safeNotification(context, notification);
            });
        } catch (error) {
            return {
                isError: true,
                content: [{ type: "text", text: `Failed to download job package: ${String(error)}` }]
            };
        }
    }
);

//=================================

// --- Schema for Acknowledge Tool ---
const acknowledgeJobSchema = z.object({
    activityId: z.string().describe("The activity/job ID to acknowledge."),
    correlationId: z.string().describe("Correlation ID returned from dequeue."),
    popReceipt: z.string().describe("PopReceipt returned from dequeue."),
    downloadLocation: z.string().url().describe("Download location returned from dequeue."),
    isDownloadFileExist: z.boolean().describe("Whether the file exists for download."),
    fileDownloadErrorMessage: z.string().optional().nullable().describe("Error message if file download failed."),
    lastDequeueDateTime: z.string().optional().nullable().describe("Last dequeue datetime.")
});

// --- Tool Definition ---
server.tool(
    "acknowledge-job",
    "Acknowledges a D365FO job execution with activity ID and dequeue details.",
    acknowledgeJobSchema.shape,
    async (
        {
            activityId,
            correlationId,
            popReceipt,
            downloadLocation,
            isDownloadFileExist,
            fileDownloadErrorMessage,
            lastDequeueDateTime
        }: z.infer<typeof acknowledgeJobSchema>,
        context: RequestHandlerExtra<ServerRequest, ServerNotification>
    ) => {
        try {
            const url = `${process.env.DYNAMICS_RESOURCE_URL}/api/connector/ack/${activityId}`;

            // Build request body (matches the Python payload)
            const body = {
                CorrelationId: correlationId,
                PopReceipt: popReceipt,
                DownloadLocation: downloadLocation,
                IsDownLoadFileExist: isDownloadFileExist,
                FileDownLoadErrorMessage: fileDownloadErrorMessage,
                LastDequeueDateTime: lastDequeueDateTime
            };

            return await callDynamicsApi('POST', url, body, async (notification) => {
                await safeNotification(context, notification);
            });
        } catch (error) {
            return {
                isError: true,
                content: [
                    {
                        type: "text",
                        text: `Failed to acknowledge job: ${String(error)}`
                    }
                ]
            };
        }
    }
);





    return server;
};
