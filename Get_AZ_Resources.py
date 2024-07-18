import pandas as pd
from azure.identity import DefaultAzureCredential
from azure.mgmt.resource import ResourceManagementClient
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime

# Initialize credentials and client
credential = DefaultAzureCredential()
subscription_ids = [
    '7f9b36c9-5325-495c-aa2a-7c7350f302a0',
    'f141e7fc-929b-4ae9-a968-94222e866042',
    'ee1c5b76-97f0-46ef-bbb0-62e6f080456b',
    '692efea0-2038-48e2-bf74-a333fef44d58',
    '7bd7cad1-cba0-45dd-8962-6efa16660e26',
    'f18b2c7d-1eb0-46d3-9ebf-1d5073e05239'
]

subscription_map = {
    '7f9b36c9-5325-495c-aa2a-7c7350f302a0': 'bmgf-eds-nonprod-00',
    'f141e7fc-929b-4ae9-a968-94222e866042': 'bmgf-eds-nonprod-traditional-00',
    'ee1c5b76-97f0-46ef-bbb0-62e6f080456b': 'bmgf-eds-prod-00',
    '692efea0-2038-48e2-bf74-a333fef44d58': 'Global Data & Analytics - NonProd',
    '7bd7cad1-cba0-45dd-8962-6efa16660e26': 'Global Data & Analytics - Prod',
    'f18b2c7d-1eb0-46d3-9ebf-1d5073e05239': 'ISS Azure PROD'
}

location_map = {
    'eastus': 'East US',
    'westus': 'West US',
    'northeurope': 'North Europe',
    'australiaeast': 'Australia East'
}

type_map = {
    'microsoft.insights/components': 'Application Insights',
    'microsoft.cache/redis': 'Azure Cache for Redis',
    'microsoft.containerregistry/registries': 'Container registry'
}

excluded_types = [
    'dell.storage/filesystems', 
    'microsoft.cdn/profiles/customdomains', 
    'microsoft.sovereign/landingzoneconfigurations', 
    'microsoft.hardwaresecuritymodules/cloudhsmclusters', 
    'microsoft.cloudtest/accounts', 
    'microsoft.cloudtest/hostedpools', 
    'microsoft.cloudtest/images', 
    'microsoft.cloudtest/pools', 
    'microsoft.compute/computefleetinstances', 
    'microsoft.compute/standbypoolinstance', 
    'microsoft.compute/virtualmachineflexinstances', 
    'microsoft.kubernetesconfiguration/extensions', 
    'microsoft.containerservice/managedclusters/microsoft.kubernetesconfiguration/extensions', 
    'microsoft.kubernetes/connectedclusters/microsoft.kubernetesconfiguration/namespaces', 
    'microsoft.containerservice/managedclusters/microsoft.kubernetesconfiguration/namespaces', 
    'microsoft.kubernetes/connectedclusters/microsoft.kubernetesconfiguration/fluxconfigurations', 
    'microsoft.containerservice/managedclusters/microsoft.kubernetesconfiguration/fluxconfigurations', 
    'microsoft.portalservices/extensions/deployments', 
    'microsoft.portalservices/extensions', 
    'microsoft.portalservices/extensions/slots', 
    'microsoft.portalservices/extensions/versions', 
    'microsoft.datacollaboration/workspaces', 
    'microsoft.deviceregistry/devices', 
    'microsoft.deviceupdate/updateaccounts/activedeployments', 
    'microsoft.deviceupdate/updateaccounts/agents', 
    'microsoft.deviceupdate/updateaccounts/deployments', 
    'microsoft.deviceupdate/updateaccounts/deviceclasses', 
    'microsoft.deviceupdate/updateaccounts/updates', 
    'microsoft.deviceupdate/updateaccounts', 
    'microsoft.devopsinfrastructure/pools', 
    'microsoft.network/dnsresolverdomainlists', 
    'microsoft.network/dnsresolverpolicies', 
    'microsoft.impact/connectors', 
    'microsoft.edgeorder/virtual_orderitems', 
    'microsoft.workloads/epicvirtualinstances', 
    'microsoft.fairfieldgardens/provisioningresources/provisioningpolicies', 
    'microsoft.fairfieldgardens/provisioningresources', 
    'microsoft.fileshares/fileshares', 
    'microsoft.healthmodel/healthmodels', 
    'microsoft.hybridcompute/arcserverwithwac', 
    'microsoft.hybridcompute/machinessovereign', 
    'microsoft.hybridcompute/machinesesu', 
    'microsoft.network/virtualhubs', 
    'microsoft.network/networkvirtualappliances', 
    'microsoft.modsimworkbench/workbenches/chambers', 
    'microsoft.modsimworkbench/workbenches/chambers/connectors', 
    'microsoft.modsimworkbench/workbenches/chambers/files', 
    'microsoft.modsimworkbench/workbenches/chambers/filerequests', 
    'microsoft.modsimworkbench/workbenches/chambers/licenses', 
    'microsoft.modsimworkbench/workbenches/chambers/storages', 
    'microsoft.modsimworkbench/workbenches/chambers/workloads', 
    'microsoft.modsimworkbench/workbenches/sharedstorages', 
    'microsoft.insights/diagnosticsettings', 
    'microsoft.network/serviceendpointpolicies', 
    'microsoft.resources/resourcegraphvisualizer', 
    'microsoft.openlogisticsplatform/workspaces', 
    'microsoft.iotoperationsmq/mq', 
    'microsoft.orbital/cloudaccessrouters', 
    'microsoft.orbital/terminals', 
    'microsoft.orbital/sdwancontrollers', 
    'microsoft.recommendationsservice/accounts/modeling', 
    'microsoft.recommendationsservice/accounts/serviceendpoints', 
    'microsoft.recoveryservicesbvtd/vaults', 
    'microsoft.recoveryservicesbvtd2/vaults', 
    'microsoft.recoveryservicesintd/vaults', 
    'microsoft.recoveryservicesintd2/vaults', 
    'microsoft.features/featureprovidernamespaces/featureconfigurations', 
    'microsoft.deploymentmanager/rollouts', 
    'microsoft.providerhub/providerregistrations', 
    'microsoft.providerhub/providerregistrations/customrollouts', 
    'microsoft.providerhub/providerregistrations/defaultrollouts', 
    'microsoft.datareplication/replicationvaults', 
    'microsoft.synapse/workspaces/sqlpools', 
    'microsoft.mission/catalogs', 
    'microsoft.mission/communities', 
    'microsoft.mission/communities/communityendpoints', 
    'microsoft.mission/enclaveconnections', 
    'microsoft.mission/virtualenclaves/enclaveendpoints', 
    'microsoft.mission/virtualenclaves/endpoints', 
    'microsoft.mission/externalconnections', 
    'microsoft.mission/internalconnections', 
    'microsoft.mission/communities/transithubs', 
    'microsoft.mission/virtualenclaves', 
    'microsoft.mission/virtualenclaves/workloads', 
    'microsoft.windowspushnotificationservices/registrations', 
    'microsoft.workloads/insights', 
    'microsoft.hanaonazure/sapmonitors', 
    'microsoft.cloudhealth/healthmodels', 
    'microsoft.connectedcache/enterprisemcccustomers/enterprisemcccachenodes', 
    'microsoft.manufacturingplatform/manufacturingdataservices', 
    'microsoft.windowsesu/multipleactivationkeys', 
    'microsoft.sql/servers/databases', 
    'microsoft.sql/servers'
]

def get_resource_group_from_id(resource_id):
    parts = resource_id.split('/')
    try:
        rg_index = parts.index('resourceGroups') + 1
        return parts[rg_index]
    except ValueError:
        return None

processed_data = []

for subscription_id in subscription_ids:
    resource_client = ResourceManagementClient(credential, subscription_id)
    resources = resource_client.resources.list()

    for item in resources:
        if item.type not in excluded_types:
            subscription_display_name = subscription_map.get(subscription_id, subscription_id)
            location_display_name = location_map.get(item.location, item.location)
            type_display_name = type_map.get(item.type, item.type)
            
            tags = item.tags if item.tags else {}
            resource_group = get_resource_group_from_id(item.id)

            processed_data.append({
                "SUBSCRIPTION_NAME": subscription_display_name,
                "SUBSCRIPTION_ID": subscription_id,
                "RESOURCE_GROUP": resource_group,
                "RESOURCE_NAME": item.name,
                "LOCATION": location_display_name,
                "TYPE": type_display_name,
                "KIND": item.kind,
                "TAG_APPLICATION": tags.get("application"),
                "TAG_OWNER": tags.get("owner"),
                "TAG_COST_CENTER": tags.get("cost-center"),
                "TAG_ENVIRONMENT": tags.get("environment"),
                "ID": item.id
            })

# Create a DataFrame
df = pd.DataFrame(processed_data)

# Sort by the 'TAG_APPLICATION' column (case-insensitive)
df = df.sort_values(by='TAG_APPLICATION', key=lambda x: x.str.lower())

# Sort by the 'SUBSCRIPTION_NAME' column (case-insensitive)
df = df.sort_values(by='SUBSCRIPTION_NAME', key=lambda x: x.str.lower())

# Generate the output file name based on the current date
current_date = datetime.now().strftime("%B_%d_%Y")
output_file = f"{current_date}_EDS_DIO_AZ_Resources_Inventory.xlsx"

# Create an Excel writer using openpyxl
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Pulled_Azure_Resources")

    # Access the workbook and the worksheet
    workbook = writer.book
    worksheet = writer.sheets["Pulled_Azure_Resources"]

    # Define the range for the table
    ref = f"A1:{chr(65 + df.shape[1] - 1)}{df.shape[0] + 1}"

    # Create a table
    table = Table(displayName="Table1", ref=ref)

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name="TableStyleMedium15", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style

    # Add the table to the worksheet
    worksheet.add_table(table)

    # Apply header style
    header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    header_font = Font(color="FFFFFF")

    for cell in worksheet["1:1"]:
        cell.fill = header_fill
        cell.font = header_font

    # Autofit column widths
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width

print(f"Data has been successfully exported to {output_file}")
