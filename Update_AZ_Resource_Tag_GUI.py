import os
import logging
from datetime import datetime
from azure.identity import DefaultAzureCredential
from azure.mgmt.resource import ResourceManagementClient
from azure.core.exceptions import ClientAuthenticationError
import tkinter as tk
from tkinter import messagebox, ttk
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
        list(client.resources.list_by_resource_group(resource_group_entry.get()))
        return client
    except ClientAuthenticationError:
        log_and_print("Authentication failed. Please re-authenticate.")
    except Exception as e:
        log_and_print(f"Error during authentication: {e}")

def log_and_print(message):
    print(message)
    logging.info(message)
    log_textbox.config(state=tk.NORMAL)
    log_textbox.insert(tk.END, message + '\n', 'message')
    log_textbox.see(tk.END)
    log_textbox.config(state=tk.DISABLED)

def configure_logging():
    log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 f"Update_AZ_Resource_Tag_{datetime.now().strftime('%Y_%m_%d_%H_%M_%S')}.log")
    logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(message)s')

def update_tags():
    global cancel_flag
    cancel_flag = False

    def run_update():
        global rollback_tags, resource
        client = authenticate_to_azure()
        if not client:
            messagebox.showerror("Authentication Error", "Failed to authenticate to Azure. Please check your credentials.")
            finalize_task("Authentication Error")
            return

        configure_logging()
        resource_group_name, resource_name = resource_group_entry.get(), resource_name_entry.get()
        tags = {k: v for k, v in {"owner": owner_tag_entry.get(), "application": application_tag_entry.get(),
                                  "environment": environment_tag_entry.get(), "cost-center": cost_center_tag_entry.get()}.items() if v}

        try:
            resource = next((r for r in client.resources.list_by_resource_group(resource_group_name) if r.name == resource_name), None)
            if resource:
                resource_id, resource_tags = resource.id, resource.tags or {}
                updated_tags = {**resource_tags, **tags}

                log_and_print(f"Current tags for resource '{resource_name}': {resource_tags}")
                if resource_tags != updated_tags:
                    log_and_print(f"Updating tags for resource: '{resource.name}'")
                    rollback_tags = resource_tags.copy()
                    if cancel_flag:
                        log_and_print("Update cancelled before execution.")
                        return

                    api_version = get_api_version(client, resource)
                    if api_version:
                        perform_update(client, resource_id, api_version, updated_tags, resource_name)
                    else:
                        log_and_print(f"Could not find API version for resource '{resource_name}'")
                else:
                    log_and_print("Tags are already up-to-date. No update required.")
                    display_tags(resource_tags)
                    messagebox.showinfo("Info", "Tags are already up-to-date. No update required.")
            else:
                log_and_print(f"Resource '{resource_name}' not found in resource group '{resource_group_name}'")
                messagebox.showerror("Error", f"Resource '{resource_name}' not found in resource group '{resource_group_name}'")

        except Exception as e:
            log_and_print(f"Error during resource retrieval or update: {e}")
            messagebox.showerror("Error", f"Error during resource retrieval or update: {e}")

        finalize_task("Task Completed")

    start_task(run_update, "Updating tags...")

def perform_update(client, resource_id, api_version, updated_tags, resource_name):
    try:
        update_operation = client.resources.begin_update_by_id(resource_id=resource_id, api_version=api_version, parameters={"tags": updated_tags})
        while not update_operation.done():
            if cancel_flag:
                log_and_print("Update cancelled during execution.")
                update_operation.cancel()
                return
            time.sleep(1)
        update_operation.result()
        log_and_print(f"Successfully updated tags for resource: '{resource_name}'\nUpdated tags: {updated_tags}")
        display_tags(updated_tags)
        populate_input_fields(updated_tags)
        messagebox.showinfo("Success", f"Successfully updated tags for resource: '{resource_name}'")
        rollback_button.config(state=tk.NORMAL)
    except Exception as e:
        log_and_print(f"Failed to update tags for resource: '{resource_name}'. Error: {str(e)}")
        rollback_update()
        messagebox.showerror("Error", f"Failed to update tags for resource: '{resource_name}'. Error: {str(e)}")

def rollback_update():
    if rollback_tags and resource:
        def run_rollback():
            client = authenticate_to_azure()
            if not client:
                messagebox.showerror("Authentication Error", "Failed to authenticate to Azure. Please check your credentials.")
                finalize_task("Authentication Error")
                return

            configure_logging()

            try:
                api_version = get_api_version(client, resource)
                if api_version:
                    log_and_print(f"Rolling back tags for resource: '{resource.name}'")
                    rollback_operation = client.resources.begin_update_by_id(resource_id=resource.id, api_version=api_version, parameters={"tags": rollback_tags})
                    while not rollback_operation.done():
                        if cancel_flag:
                            log_and_print("Rollback cancelled during execution.")
                            rollback_operation.cancel()
                            return
                        time.sleep(1)
                    rollback_operation.result()
                    log_and_print(f"Successfully rolled back tags for resource: '{resource.name}'")
                    display_tags(rollback_tags)
                    populate_input_fields(rollback_tags)
                    messagebox.showinfo("Success", f"Successfully rolled back tags for resource: '{resource.name}'")
                else:
                    log_and_print("Failed to get API version. Cannot rollback.")
                    messagebox.showerror("Error", "Failed to get API version. Cannot rollback.")
            except Exception as e:
                log_and_print(f"Failed to rollback tags: {str(e)}")
                messagebox.showerror("Error", f"Failed to rollback tags: {str(e)}")

            finalize_task("Rollback Completed")

        start_task(run_rollback, "Rolling back tags...")

def pull_resource_tags():
    global cancel_flag
    cancel_flag = False

    def run_pull():
        client = authenticate_to_azure()
        if not client:
            messagebox.showerror("Authentication Error", "Failed to authenticate to Azure. Please check your credentials.")
            finalize_task("Authentication Error")
            return

        configure_logging()
        resource_group_name, resource_name = resource_group_entry.get(), resource_name_entry.get()

        try:
            resource = next((r for r in client.resources.list_by_resource_group(resource_group_name) if r.name == resource_name), None)
            if resource:
                resource_tags = resource.tags or {}
                log_and_print(f"Current tags for resource '{resource_name}': {resource_tags}")
                display_tags(resource_tags)
                populate_input_fields(resource_tags)
                messagebox.showinfo("Success", f"Successfully pulled tags for resource: '{resource.name}'")
            else:
                log_and_print(f"Resource '{resource_name}' not found in resource group '{resource_group_name}'")
                messagebox.showerror("Error", f"Resource '{resource_name}' not found in resource group '{resource_group_name}'")
        except Exception as e:
            log_and_print(f"Error during resource retrieval: {e}")
            messagebox.showerror("Error", f"Error during resource retrieval: {e}")

        finalize_task("Pull Task Completed")

    start_task(run_pull, "Pulling tags...")

def get_api_version(client, resource):
    try:
        namespace, r_type = resource.type.split('/')
        provider = client.providers.get(namespace)
        r_type_obj = next((rt for rt in provider.resource_types if rt.resource_type == r_type), None)
        return r_type_obj.api_versions[0] if r_type_obj else None
    except Exception as e:
        log_and_print(f"Error getting API version: {e}")
        return None

def start_task(target, progress_text):
    disable_buttons(during_update=True)
    # cancel_button.config(state=tk.NORMAL)  # Commented out
    progress_bar.start()
    progress_label.config(text=progress_text)
    threading.Thread(target=target).start()

def finalize_task(status_text):
    enable_buttons()
    progress_bar.stop()
    progress_label.config(text=status_text)
    log_and_print("Task Completed\n*************************************************************************************************************************")

def clear_inputs():
    for entry in [subscription_id_entry, resource_group_entry, resource_name_entry, owner_tag_entry,
                  application_tag_entry, environment_tag_entry, cost_center_tag_entry]:
        entry.delete(0, tk.END)
    log_textbox.config(state=tk.NORMAL)
    log_textbox.delete(1.0, tk.END)
    log_textbox.config(state=tk.DISABLED)
    disable_buttons()

def close_application():
    root.destroy()

def disable_buttons(during_update=False):
    state = tk.DISABLED
    update_button.config(state=state)
    pull_button.config(state=state)
    clear_button.config(state=state)
    rollback_button.config(state=state)
    # cancel_button.config(state=tk.NORMAL if during_update else state)  # Commented out
    close_button.config(state=tk.DISABLED if during_update else tk.NORMAL)

def enable_buttons():
    state = tk.NORMAL if all([subscription_id_entry.get(), resource_group_entry.get(), resource_name_entry.get()]) else tk.DISABLED
    update_button.config(state=state)
    pull_button.config(state=state)
    clear_button.config(state=state)
    rollback_button.config(state=tk.NORMAL if rollback_tags else state)
    # cancel_button.config(state=tk.DISABLED)  # Commented out
    close_button.config(state=tk.NORMAL)

def disable_close():
    pass

def display_tags(tags):
    log_textbox.config(state=tk.NORMAL)
    log_textbox.insert(tk.END, "Tags:\n", 'header')
    for key, value in tags.items():
        log_textbox.insert(tk.END, f"{key}: {value}\n", 'message')
    log_textbox.see(tk.END)
    log_textbox.config(state=tk.DISABLED)

def populate_input_fields(tags):
    for entry, key in [(owner_tag_entry, "owner"), (application_tag_entry, "application"), 
                       (environment_tag_entry, "environment"), (cost_center_tag_entry, "cost-center")]:
        entry.delete(0, tk.END)
        entry.insert(0, tags.get(key, ""))

def validate_input_fields(event=None):
    enable_buttons() if all([subscription_id_entry.get(), resource_group_entry.get(), resource_name_entry.get()]) else disable_buttons()

# Create the main window
root = tk.Tk()
root.title("Azure Resource Tag Updater")
root.geometry("680x500")

root.update_idletasks()
width, height = root.winfo_width(), root.winfo_height()
x, y = (root.winfo_screenwidth() // 2) - (width // 2), (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f'{width}x{height}+{x}+{y}')
root.protocol("WM_DELETE_WINDOW", disable_close)

def load_icon(path):
    return ImageTk.PhotoImage(Image.open(path).resize((20, 20), Image.LANCZOS))

icons = {name: load_icon(os.path.join("icons", f"{name}.jpg")) for name in ["update", "clear", "update", "rollback", "close", "pull"]}

def create_label_and_entry(row, label_text, variable):
    tk.Label(root, text=label_text).grid(row=row, column=0, padx=5, pady=2, sticky="e")
    entry = tk.Entry(root, textvariable=variable)
    entry.grid(row=row, column=1, padx=5, pady=2, sticky="w")
    return entry

variables = {name: tk.StringVar() for name in ["subscription_id", "resource_group", "resource_name", "owner_tag", "application_tag", "environment_tag", "cost_center_tag"]}
entries = [create_label_and_entry(i, label, variables[name]) for i, (name, label) in enumerate([
    ("subscription_id", "Subscription ID:"), ("resource_group", "Resource Group Name:"), ("resource_name", "Resource Name:"),
    ("owner_tag", "Owner Tag Value:"), ("application_tag", "Application Tag Value:"), ("environment_tag", "Environment Tag Value:"),
    ("cost_center_tag", "Cost-center Tag Value:")])]
subscription_id_entry, resource_group_entry, resource_name_entry, owner_tag_entry, application_tag_entry, environment_tag_entry, cost_center_tag_entry = entries

for entry in [subscription_id_entry, resource_group_entry, resource_name_entry]:
    entry.bind("<KeyRelease>", validate_input_fields)

for widget in root.winfo_children():
    widget.grid_configure(sticky="ew")

button_frame = tk.Frame(root)
button_frame.grid(row=7, column=0, columnspan=2, pady=(20, 5))

button_config = [("Update Tags", update_tags, "update"), ("Pull Resource Tags", pull_resource_tags, "pull"), 
                 ("Clear Inputs", clear_inputs, "clear"), ("Rollback Tags", rollback_update, "rollback"), ("Close", close_application, "close")]
buttons = {text: tk.Button(button_frame, text=text, command=command, state=tk.DISABLED if text != "Close" else tk.NORMAL, image=icons[icon], compound=tk.LEFT) 
           for text, command, icon in button_config}
for i, button in enumerate(buttons.values()):
    button.grid(row=0, column=i, padx=5)

update_button, pull_button, clear_button, rollback_button, close_button = buttons.values()

progress_bar = ttk.Progressbar(root, mode="indeterminate")
progress_bar.grid(row=8, column=0, columnspan=2, pady=10, sticky="we")
progress_label = tk.Label(root, text="")
progress_label.grid(row=9, column=0, columnspan=2)

log_textbox_frame = tk.Frame(root)
log_textbox_frame.grid(row=10, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
log_textbox_scrollbar = tk.Scrollbar(log_textbox_frame)
log_textbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_textbox = tk.Text(log_textbox_frame, height=20, width=80, yscrollcommand=log_textbox_scrollbar.set, wrap=tk.WORD, state=tk.DISABLED, insertontime=0, insertofftime=0)
log_textbox.pack(expand=True, fill=tk.BOTH)
log_textbox_scrollbar.config(command=log_textbox.yview)
log_textbox.config(font=("Helvetica", 10), padx=10, pady=10, bg="#f0f0f0")

log_textbox.tag_configure('header', font=('Helvetica', 12, 'bold'))
log_textbox.tag_configure('message', font=('Helvetica', 10))

root.grid_rowconfigure(10, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

root.mainloop()
