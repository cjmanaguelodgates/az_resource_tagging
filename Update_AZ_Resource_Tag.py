import os
import logging
from datetime import datetime
from azure.identity import DefaultAzureCredential
from azure.mgmt.resource import ResourceManagementClient

# Parameters
subscription_input = input("Enter the subscription ID: ")
resource_group_name = input("Enter the resource group name: ")
resource_name = input("Enter the resource name: ")
owner_tag = input("Enter the owner tag value: ")
application_tag = input("Enter the application tag value: ")
environment_tag = input("Enter the environment tag value: ")
cost_center_tag = input("Enter the cost-center tag value: ")

tags = {
    "owner": owner_tag,
    "application": application_tag,
    "environment": environment_tag,
    "cost-center": cost_center_tag
}

# Set the default log directory to the user's home directory
log_directory = os.path.join(os.path.expanduser("~"), "ResourceTagLogs")
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

# Set up logging
log_file_path = os.path.join(log_directory, f"ResourceTagUpdate_{resource_group_name}_{datetime.now().strftime('%Y%m%d')}.log")
logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(message)s')

def log_and_print(message):
    print(message)
    logging.info(message)

# Authenticate to Azure
credential = DefaultAzureCredential()
client = ResourceManagementClient(credential, subscription_input)

# Get the specific resource in the resource group
resources = client.resources.list_by_resource_group(resource_group_name)
resource = next((r for r in resources if r.name == resource_name), None)

if resource:
    resource_id = resource.id
    resource_tags = resource.tags if resource.tags else {}
    updated_tags = {**resource_tags, **tags}

    # Get the appropriate API version for the resource type
    resource_provider_namespace = resource.type.split('/')[0]
    resource_type = resource.type.split('/')[1]
    provider = client.providers.get(resource_provider_namespace)
    resource_type_obj = next((rt for rt in provider.resource_types if rt.resource_type == resource_type), None)
    if resource_type_obj:
        api_version = resource_type_obj.api_versions[0]  # Use the latest API version
    else:
        log_and_print(f"Could not find resource type '{resource_type}' for provider '{resource_provider_namespace}'")
        api_version = None

    if api_version and resource_tags != updated_tags:
        log_and_print(f"Updating tags for resource: '{resource.name}' in '{resource_group_name}' from '{subscription_input}' subscription.")
        try:
            client.resources.begin_update_by_id(
                resource_id=resource_id,
                api_version=api_version,
                parameters={"tags": updated_tags}
            ).result()
            updated_tags_diff = {key: updated_tags[key] for key in tags.keys() if updated_tags[key] != resource_tags.get(key)}
            log_and_print(f"Successfully updated tags for resource: '{resource.name}'. Updated tags: {updated_tags_diff}")
        except Exception as e:
            log_and_print(f"Failed to update tags for resource: '{resource.name}'. Error: {str(e)}")
else:
    log_and_print(f"Resource '{resource_name}' not found in resource group '{resource_group_name}'.")

log_and_print("Task Completed")
