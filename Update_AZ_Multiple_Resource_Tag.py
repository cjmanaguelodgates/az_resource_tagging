import os
import logging
import pandas as pd
from datetime import datetime
from azure.identity import DefaultAzureCredential
from azure.mgmt.resource import ResourceManagementClient
from azure.mgmt.resource.resources.models import GenericResource, Sku

# Set the default log directory to the user's home directory
log_directory = os.path.join(os.path.expanduser("~"), "ResourceTagLogs")
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

# Set up logging
log_file_path = os.path.join(log_directory, f"ResourceTagUpdate_{datetime.now().strftime('%Y%m%d')}.log")
logging.basicConfig(filename=log_file_path, level=logging.DEBUG, format='%(asctime)s - %(message)s')

def log_and_print(message):
    print(message)
    logging.debug(message)

# Authenticate to Azure
credential = DefaultAzureCredential()

# Load the Excel file
excel_file = 'resource_tags.xlsx'
df = pd.read_excel(excel_file)

successful_updates_count = 0

# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    subscription_id = row['subscription_id']
    resource_group_name = row['resource_group_name']
    resource_name = row['resource_name']
    resource_type = row.get('resource_type', None)  # Handle missing column gracefully
    resource_type = str(resource_type) if not pd.isna(resource_type) else None  # Ensure resource_type is a string or None
    owner_tag = row['owner_tag']
    application_tag = row['application_tag']
    environment_tag = row['environment_tag']
    cost_center_tag = row['cost_center_tag']

    tags = {
        "owner": owner_tag,
        "application": application_tag,
        "environment": environment_tag,
        "cost-center": cost_center_tag
    }

    # Initialize the Resource Management Client
    client = ResourceManagementClient(credential, subscription_id)

    # Get the resource by name, and type if provided
    try:
        resources = client.resources.list_by_resource_group(resource_group_name)
        if resource_type:
            resource = next((res for res in resources if res.name == resource_name and res.type == resource_type), None)
        else:
            resource = next((res for res in resources if res.name == resource_name), None)

        if resource is None:
            log_and_print(f"Resource '{resource_name}' {'of type ' + resource_type if resource_type else ''} not found in resource group '{resource_group_name}'")
            continue
    except Exception as e:
        log_and_print(f"Failed to get resource: '{resource_name}' {'of type ' + resource_type if resource_type else ''} in resource group '{resource_group_name}'. Error: {str(e)}")
        continue

    resource_tags = resource.tags if resource.tags else {}
    updated_tags = {**resource_tags, **tags}

    if resource_tags != updated_tags:
        log_and_print(f"Updating tags for resource: '{resource_name}' in '{resource_group_name}' from '{subscription_id}' subscription.")
        try:
            # Get the API version for the resource
            provider_namespace, resource_type_name = resource.type.split('/', 1)
            resource_provider = client.providers.get(provider_namespace)
            resource_type_info = next((rt for rt in resource_provider.resource_types if rt.resource_type == resource_type_name), None)
            if resource_type_info is None:
                log_and_print(f"Resource type '{resource_type_name}' not found for provider '{provider_namespace}'")
                continue
            api_version = resource_type_info.api_versions[0]

            # Ensure properties is not None
            resource_properties = resource.properties if resource.properties else {}

            # Ensure SKU is not None if it is required
            resource_sku = resource.sku if resource.sku else None

            # Create a GenericResource object with the necessary parameters
            generic_resource = GenericResource(
                location=resource.location,
                tags=updated_tags,
                properties=resource_properties,
                sku=resource_sku
            )

            # Log detailed information
            log_and_print(f"Resource ID: {resource.id}")
            log_and_print(f"API Version: {api_version}")
            log_and_print(f"Generic Resource: {generic_resource}")

            # Attempt to update the resource with new tags
            response = client.resources.begin_update_by_id(
                resource.id,
                api_version,
                parameters=generic_resource
            ).result()

            if response:
                updated_tags_diff = {key: updated_tags[key] for key in tags.keys() if updated_tags[key] != resource_tags.get(key)}
                log_and_print(f"Successfully updated tags for resource: '{resource_name}'. Updated tags: {updated_tags_diff}")
                successful_updates_count += 1
            else:
                log_and_print(f"Failed to update tags for resource: '{resource_name}'. No response from Azure.")
        except Exception as e:
            log_and_print(f"Failed to update tags for resource: '{resource_name}'. Error: {str(e)}")
    else:
        log_and_print(f"Tags for resource '{resource_name}' are already up to date.")

log_and_print(f"Task Completed")
log_and_print(f"Number of resources with successful tag updates: {successful_updates_count}")
