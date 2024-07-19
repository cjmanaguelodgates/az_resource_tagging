import os
import logging
from datetime import datetime
from azure.identity import DefaultAzureCredential
from azure.mgmt.resource import ResourceManagementClient
from azure.core.exceptions import ClientAuthenticationError
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import threading
import time
from PIL import Image, ImageTk

# Global variables
cancel_flag = False
rollback_tags = None
resource = None

def authenticate_to_azure():
    try:
        credential = DefaultAzureCredential()
        client = ResourceManagementClient(credential, subscription_id_entry.get())
        # Perform a test call to ensure authentication is valid
        list(client.resources.list_by_resource_group(resource_group_entry.get()))
        return client
    except ClientAuthenticationError:
        log_and_print("Authentication failed. Please re-authenticate.")
        return None
    except Exception as e:
        log_and_print(f"Error during authentication: {e}")
        return None

def log_and_print(message):
    print(message)
    logging.info(message)
    log_textbox.insert(tk.END, message + '\n')
    log_textbox.see(tk.END)

def update_tags():
    global cancel_flag
    cancel_flag = False

    def run_update():
        global rollback_tags
        global resource
        client = authenticate_to_azure()
        if client is None:
            messagebox.showerror("Authentication Error", "Failed to authenticate to Azure. Please check your credentials.")
            enable_buttons()
            progress_bar.stop()
            progress_label.config(text="Authentication Error")
            return

        subscription_input = subscription_id_entry.get()
        resource_group_name = resource_group_entry.get()
        resource_name = resource_name_entry.get()
        owner_tag = owner_tag_entry.get()
        application_tag = application_tag_entry.get()
        environment_tag = environment_tag_entry.get()
        cost_center_tag = cost_center_tag_entry.get()

        tags = {
            "owner": owner_tag if owner_tag else None,
            "application": application_tag if application_tag else None,
            "environment": environment_tag if environment_tag else None,
            "cost-center": cost_center_tag if cost_center_tag else None
        }

        # Remove None values from tags
        tags = {k: v for k, v in tags.items() if v is not None}

        # Set the log directory to the script file's directory
        script_directory = os.path.dirname(os.path.abspath(__file__))
        log_directory = script_directory

        # Set up logging
        log_file_path = os.path.join(log_directory, f"ResourceTagUpdate_{resource_group_name}_{datetime.now().strftime('%Y%m%d')}.log")
        logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(message)s')

        def log_tags(resource_tags, stage):
            log_and_print(f"{stage} tags for resource '{resource_name}': {resource_tags}")

        # Get the specific resource in the resource group
        try:
            resources = client.resources.list_by_resource_group(resource_group_name)
            resource = next((r for r in resources if r.name == resource_name), None)

            if resource:
                resource_id = resource.id
                resource_tags = resource.tags if resource.tags else {}
                updated_tags = {**resource_tags, **tags}

                # Log existing tags
                log_tags(resource_tags, "Current")

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
                        # Backup current tags for rollback
                        rollback_tags = resource_tags.copy()

                        # Check if the cancel flag is set before making the update
                        if cancel_flag:
                            log_and_print("Update cancelled before execution.")
                            return

                        update_operation = client.resources.begin_update_by_id(
                            resource_id=resource_id,
                            api_version=api_version,
                            parameters={"tags": updated_tags}
                        )

                        # Wait for the update operation to complete, periodically checking the cancel flag
                        while not update_operation.done():
                            if cancel_flag:
                                log_and_print("Update cancelled during execution.")
                                update_operation.cancel()
                                return
                            time.sleep(1)

                        update_operation.result()  # Ensure the operation is complete

                        updated_tags_diff = {key: updated_tags[key] for key in tags.keys() if updated_tags[key] != resource_tags.get(key)}
                        log_and_print(f"Successfully updated tags for resource: '{resource.name}'. Updated tags: {updated_tags_diff}")
                        # Log updated tags
                        log_tags(updated_tags, "Updated")
                        display_tags(updated_tags)
                        populate_input_fields(updated_tags)  # Populate input fields with updated tags
                        messagebox.showinfo("Success", f"Successfully updated tags for resource: '{resource.name}'.")
                        rollback_button.config(state=tk.NORMAL)  # Enable rollback button since there are tags processed
                    except Exception as e:
                        log_and_print(f"Failed to update tags for resource: '{resource.name}'. Error: {str(e)}")
                        rollback_update()
                        messagebox.showerror("Error", f"Failed to update tags for resource: '{resource.name}'. Error: {str(e)}")
                else:
                    log_and_print("Tags are already up-to-date or no valid tag values provided. No update required.")
                    display_tags(resource_tags)
                    messagebox.showinfo("Info", "Tags are already up-to-date or no valid tag values provided. No update required.")
            else:
                log_and_print(f"Resource '{resource_name}' not found in resource group '{resource_group_name}'.")
                messagebox.showerror("Error", f"Resource '{resource_name}' not found in resource group '{resource_group_name}'.")

        except ClientAuthenticationError:
            log_and_print("Authentication failed. Please re-authenticate.")
            messagebox.showerror("Authentication Error", "Failed to authenticate to Azure. Please check your credentials.")
        except Exception as e:
            log_and_print(f"Error during resource retrieval or update: {e}")
            messagebox.showerror("Error", f"Error during resource retrieval or update: {e}")

        log_and_print("Task Completed")
        enable_buttons()
        progress_bar.stop()
        progress_label.config(text="Task Completed")

    disable_buttons(during_update=True)
    cancel_button.config(state=tk.NORMAL)
    progress_bar.start()
    progress_label.config(text="Updating tags...")
    threading.Thread(target=run_update).start()

def rollback_update():
    if rollback_tags is not None and resource is not None:
        def run_rollback():
            try:
                client = authenticate_to_azure()
                if client is None:
                    messagebox.showerror("Authentication Error", "Failed to authenticate to Azure. Please check your credentials.")
                    enable_buttons()
                    progress_bar.stop()
                    progress_label.config(text="Authentication Error")
                    return

                subscription_input = subscription_id_entry.get()
                resource_group_name = resource_group_entry.get()
                resource_name = resource_name_entry.get()

                # Set the log directory to the script file's directory
                script_directory = os.path.dirname(os.path.abspath(__file__))
                log_directory = script_directory

                # Set up logging
                log_file_path = os.path.join(log_directory, f"ResourceTagUpdate_{resource_group_name}_{datetime.now().strftime('%Y%m%d')}.log")
                logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(message)s')

                def log_tags(resource_tags, stage):
                    log_and_print(f"{stage} tags for resource '{resource_name}': {resource_tags}")

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

                if api_version:
                    log_and_print(f"Rolling back tags for resource: '{resource.name}' in '{resource_group_name}' from '{subscription_input}' subscription.")
                    try:
                        rollback_operation = client.resources.begin_update_by_id(
                            resource_id=resource.id,
                            api_version=api_version,
                            parameters={"tags": rollback_tags}
                        )

                        # Wait for the rollback operation to complete, periodically checking the cancel flag
                        while not rollback_operation.done():
                            if cancel_flag:
                                log_and_print("Rollback cancelled during execution.")
                                rollback_operation.cancel()
                                return
                            time.sleep(1)

                        rollback_operation.result()  # Ensure the operation is complete

                        log_and_print(f"Successfully rolled back tags for resource: '{resource.name}'.")
                        # Log rolled back tags
                        log_tags(rollback_tags, "Rolled back")
                        display_tags(rollback_tags)
                        populate_input_fields(rollback_tags)  # Populate input fields with rolled-back tags
                        messagebox.showinfo("Success", f"Successfully rolled back tags for resource: '{resource.name}'.")
                    except Exception as rollback_e:
                        log_and_print(f"Failed to rollback tags for resource: '{resource.name}'. Error: {str(rollback_e)}")
                        messagebox.showerror("Error", f"Failed to rollback tags for resource: '{resource.name}'. Error: {str(rollback_e)}")
                else:
                    log_and_print("Failed to get API version. Cannot rollback.")
                    messagebox.showerror("Error", "Failed to get API version. Cannot rollback.")
            except ClientAuthenticationError:
                log_and_print("Authentication failed. Please re-authenticate.")
                messagebox.showerror("Authentication Error", "Failed to authenticate to Azure. Please check your credentials.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to rollback tags: {str(e)}")

            log_and_print("Rollback Task Completed")
            enable_buttons()
            progress_bar.stop()
            progress_label.config(text="Rollback Completed")

        disable_buttons(during_update=True)
        cancel_button.config(state=tk.DISABLED)
        progress_bar.start()
        progress_label.config(text="Rolling back tags...")
        threading.Thread(target=run_rollback).start()

def cancel_update():
    global cancel_flag
    cancel_flag = True
    enable_buttons()
    progress_bar.stop()
    progress_label.config(text="Update Cancelled")
    messagebox.showinfo("Info", "Update cancelled.")

def clear_inputs():
    # Clear all the input fields
    subscription_id_entry.delete(0, tk.END)
    resource_group_entry.delete(0, tk.END)
    resource_name_entry.delete(0, tk.END)
    owner_tag_entry.delete(0, tk.END)
    application_tag_entry.delete(0, tk.END)
    environment_tag_entry.delete(0, tk.END)
    cost_center_tag_entry.delete(0, tk.END)

    # Clear the log textbox
    log_textbox.delete(1.0, tk.END)

    # Disable all buttons
    disable_buttons()

def close_application():
    root.destroy()

def disable_buttons(during_update=False):
    update_button.config(state=tk.DISABLED)
    clear_button.config(state=tk.DISABLED)
    rollback_button.config(state=tk.DISABLED)
    cancel_button.config(state=tk.DISABLED if not during_update else tk.NORMAL)
    close_button.config(state=tk.DISABLED if during_update else tk.NORMAL)

def enable_buttons():
    if subscription_id_entry.get() and resource_group_entry.get() and resource_name_entry.get():
        update_button.config(state=tk.NORMAL)
        clear_button.config(state=tk.NORMAL)
        cancel_button.config(state=tk.DISABLED)  # Cancel button should be enabled only when updating
        rollback_button.config(state=tk.NORMAL if rollback_tags else tk.DISABLED)  # Rollback button enabled based on tags processed
        close_button.config(state=tk.NORMAL)
    else:
        update_button.config(state=tk.DISABLED)
        clear_button.config(state=tk.DISABLED)
        rollback_button.config(state=tk.DISABLED)
        cancel_button.config(state=tk.DISABLED)
        close_button.config(state=tk.NORMAL)

def disable_close():
    pass

def display_tags(tags):
    log_textbox.insert(tk.END, "Tags:\n")
    for key, value in tags.items():
        log_textbox.insert(tk.END, f"{key}: {value}\n")
    log_textbox.see(tk.END)

def populate_input_fields(tags):
    owner_tag_entry.delete(0, tk.END)
    owner_tag_entry.insert(0, tags.get("owner", ""))
    application_tag_entry.delete(0, tk.END)
    application_tag_entry.insert(0, tags.get("application", ""))
    environment_tag_entry.delete(0, tk.END)
    environment_tag_entry.insert(0, tags.get("environment", ""))
    cost_center_tag_entry.delete(0, tk.END)
    cost_center_tag_entry.insert(0, tags.get("cost-center", ""))

def validate_input_fields(event=None):
    if subscription_id_entry.get() and resource_group_entry.get() and resource_name_entry.get():
        enable_buttons()
    else:
        disable_buttons()

# Create the main window
root = tk.Tk()
root.title("Azure Resource Tag Updater")
root.geometry("600x500")

# Center the window on the screen
root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f'{width}x{height}+{x}+{y}')

# Remove window close functionality
root.protocol("WM_DELETE_WINDOW", disable_close)

# Load and resize button icons
def load_icon(path):
    img = Image.open(path)
    img = img.resize((20, 20), Image.LANCZOS)
    return ImageTk.PhotoImage(img)

update_icon = load_icon("update.jpg")
clear_icon = load_icon("clear.jpg")
cancel_icon = load_icon("update.jpg")
rollback_icon = load_icon("rollback.jpg")
close_icon = load_icon("close.jpg")

# Create and place the input fields and labels
tk.Label(root, text="Subscription ID:").grid(row=0, column=0, padx=10, pady=5)
subscription_id_entry = tk.Entry(root)
subscription_id_entry.grid(row=0, column=1, padx=10, pady=5)
subscription_id_entry.bind("<KeyRelease>", validate_input_fields)

tk.Label(root, text="Resource Group Name:").grid(row=1, column=0, padx=10, pady=5)
resource_group_entry = tk.Entry(root)
resource_group_entry.grid(row=1, column=1, padx=10, pady=5)
resource_group_entry.bind("<KeyRelease>", validate_input_fields)

tk.Label(root, text="Resource Name:").grid(row=2, column=0, padx=10, pady=5)
resource_name_entry = tk.Entry(root)
resource_name_entry.grid(row=2, column=1, padx=10, pady=5)
resource_name_entry.bind("<KeyRelease>", validate_input_fields)

tk.Label(root, text="Owner Tag Value:").grid(row=3, column=0, padx=10, pady=5)
owner_tag_entry = tk.Entry(root)
owner_tag_entry.grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="Application Tag Value:").grid(row=4, column=0, padx=10, pady=5)
application_tag_entry = tk.Entry(root)
application_tag_entry.grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="Environment Tag Value:").grid(row=5, column=0, padx=10, pady=5)
environment_tag_entry = tk.Entry(root)
environment_tag_entry.grid(row=5, column=1, padx=10, pady=5)

tk.Label(root, text="Cost-center Tag Value:").grid(row=6, column=0, padx=10, pady=5)
cost_center_tag_entry = tk.Entry(root)
cost_center_tag_entry.grid(row=6, column=1, padx=10, pady=5)

# Create and place the buttons horizontally
button_frame = tk.Frame(root)
button_frame.grid(row=7, column=0, columnspan=2, pady=5)

update_button = tk.Button(button_frame, text="Update Tags", command=update_tags, state=tk.DISABLED, image=update_icon, compound=tk.LEFT)
update_button.grid(row=0, column=0, padx=5)

cancel_button = tk.Button(button_frame, text="Cancel Update", command=cancel_update, state=tk.DISABLED, image=cancel_icon, compound=tk.LEFT)
cancel_button.grid(row=0, column=1, padx=5)

clear_button = tk.Button(button_frame, text="Clear Inputs", command=clear_inputs, state=tk.DISABLED, image=clear_icon, compound=tk.LEFT)
clear_button.grid(row=0, column=2, padx=5)

rollback_button = tk.Button(button_frame, text="Rollback Tags", command=rollback_update, state=tk.DISABLED, image=rollback_icon, compound=tk.LEFT)
rollback_button.grid(row=0, column=3, padx=5)

close_button = tk.Button(button_frame, text="Close", command=close_application, state=tk.NORMAL, image=close_icon, compound=tk.LEFT)
close_button.grid(row=0, column=4, padx=5)

# Create and place the progress bar
progress_bar = ttk.Progressbar(root, mode="determinate")
progress_bar.grid(row=8, column=0, columnspan=2, pady=10, sticky="we")

progress_label = tk.Label(root, text="")
progress_label.grid(row=9, column=0, columnspan=2)

# Create and place the scrollable textbox for displaying logs and tags
log_textbox_frame = tk.Frame(root)
log_textbox_frame.grid(row=10, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
log_textbox_scrollbar = tk.Scrollbar(log_textbox_frame)
log_textbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_textbox = tk.Text(log_textbox_frame, height=20, width=80, yscrollcommand=log_textbox_scrollbar.set)
log_textbox.pack(expand=True, fill=tk.BOTH)
log_textbox_scrollbar.config(command=log_textbox.yview)

# Make the window resizable
root.grid_rowconfigure(10, weight=1)
root.grid_columnconfigure(1, weight=1)

# Run the GUI loop
root.mainloop()
