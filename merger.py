# =============================================================================
# IMPORTS
# =============================================================================
from tkinter import simpledialog, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.dialogs import MessageDialog, Messagebox
import os, platform, subprocess


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def center_window(window):
    """Centers the given window on the screen."""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


def show_exit_confirmation(callback_function):
    """
    Shows exit confirmation dialog and handles user's choice.
    If user chooses 'No', calls the provided callback function.
    """
    exit_dialog = MessageDialog(
        title="Mail Merge",
        message="Do you want to exit?",
        buttons=["Yes", "No"]
    )
    exit_dialog.show()
    user_choice = exit_dialog.result

    if user_choice == "Yes":
        app_window.quit()
        app_window.destroy()
        exit()
    else:
        callback_function()


# =============================================================================
# NAME INPUT FUNCTIONS
# =============================================================================

def name_collection_method():
    """Ask user for name input method"""
    global input_method_choice
    name_input_method_dialog = MessageDialog(
        title="Mail Merge",
        message=name_input_method_prompt,
        buttons=["Manual Insert", "Text File"],
        parent= app_window
    )
    name_input_method_dialog.show()
    input_method_choice = name_input_method_dialog.result

    if input_method_choice is None:
        show_exit_confirmation(name_collection_method)


def collect_names_manually():
    """
    Collects names from user through dialog boxes.
    Validates that each name contains only alphabetic characters, spaces, and hyphens.
    """
    global recipient_names, name_prompt, total_recipients

    for _ in range(total_recipients - len(recipient_names)):
        while True:
            entered_name = simpledialog.askstring(title="Name", prompt=name_prompt)

            # Handle dialog cancellation
            if entered_name is None:
                show_exit_confirmation(collect_names_manually)
                break

            # Validate name contains only letters, spaces, and hyphens
            if entered_name.replace(" ", "").replace("-", "").isalpha():
                name_prompt = "Please enter the name:"
                recipient_names.append(entered_name.title())
                break
            else:
                name_prompt = "Please enter a valid name:"


def confirm_names_and_proceed():
    """
    Shows collected names to user for confirmation.
    Allows re-entry or continuation to file selection.
    Returns user's file input choice (Browse or Type).
    """
    global names_confirmation_dialog, app_window, recipient_names

    names_confirmation_dialog.show()
    confirmation_choice = names_confirmation_dialog.result

    if confirmation_choice == "Continue":
        return confirmation_choice

    elif confirmation_choice == "Re-enter Names":
        recipient_names = []
        collect_names_manually()
        confirm_names_and_proceed()
    else:
        show_exit_confirmation(confirm_names_and_proceed)


def prompt_recipient_count():
    """Prompts user to enter the number of recipients."""
    global total_recipients
    total_recipients = simpledialog.askinteger(
        title="No. of People",
        prompt=recipient_count_prompt,
        minvalue= 2,
        parent= app_window
    )


    # Handle cancellation of recipient count dialog
    if total_recipients is None:
        show_exit_confirmation(prompt_recipient_count)

def name_text_file_processing():
    # Get names file from user
    global recipient_names
    names_file_location = filedialog.askopenfilename(
        title="Select the names text file",
        filetypes=(("Text files", "*.txt"),)
    )

    with open(names_file_location) as names_file:
        names_content = names_file.read()

    # Validate comma separation
    if ", " not in names_content:
        Messagebox.show_warning(
            title="Mail Merge",
            message="Please enter the names separated by comma."
        )
        name_text_file_processing()

    recipient_names = names_content.split(", ")

# =============================================================================
# LETTER CONTENT FUNCTIONS
# =============================================================================

def show_placeholder_instructions():
    """Displays instructions about using the [name] placeholder."""
    Messagebox.show_info(
        title="Mail Merge",
        message="Please make sure that in your file type '[name]' where you want each person's name to go."
    )


def generate_personalized_letters():
    """
    Creates personalized letter files for each recipient.
    Replaces [name] placeholder with actual names and saves files.
    """
    global recipient_names, letter_content

    Messagebox.show_info(
        title="Mail Merge",
        message="Now please select the folder where you want to save the letters."
    )

    output_directory = filedialog.askdirectory(
        title="Where do you want to save the mails?"
    )

    # Create individual letter file for each recipient
    for recipient in recipient_names:
        with open(f"{output_directory}/{recipient}'s Mail.txt", "w") as output_file:
            personalized_content = letter_content.replace("[name]", recipient)
            output_file.write(personalized_content)

    Messagebox.show_info(
        title="Mail Merge",
        message="Your mail merge was successful, congratulations!"
    )

    # Open mail folder
    if platform.system() == "Windows":
        os.startfile(output_directory)
    elif platform.system() == "Darwin":  # macOS
        subprocess.call(["open", output_directory])
    else:
        subprocess.call(["xdg-open", output_directory])


def process_letter_content():
    """
    Handles letter content input based on user's choice (Browse or Type).
    Validates placeholder presence and triggers letter generation.
    """
    global file_input_choice, letter_content, is_placeholder_not_present

    file_input_dialog = MessageDialog(
        title="Mail Merge",
        message="Do you want to browse the file or paste the letter content?",
        buttons=["Browse", "Type Letter Content"]
    )
    file_input_dialog.show()
    file_input_choice = file_input_dialog.result

    if file_input_dialog.result is None:
        show_exit_confirmation(process_letter_content)

    if file_input_choice == "Browse":
        show_placeholder_instructions()

        # Get letter file from user
        letter_file_location = filedialog.askopenfilename(
            title="Select your mail",
            filetypes=(("Text files", "*.txt"),)
        )

        if not letter_file_location:
            show_exit_confirmation(process_letter_content)

        with open(letter_file_location) as letter_file:
            letter_content = letter_file.read()

            # Validate [name] placeholder exists
            name_place = letter_content.find("[name]")
            if name_place < 0:
                Messagebox.show_warning(
                    title="Mail Merge",
                    message="Warning: Placeholder '[name]' is not in the letter body"
                )
                is_placeholder_not_present = True
            else:
                is_placeholder_not_present = False

            if is_placeholder_not_present:
                # Re-prompt for file input method
                process_letter_content()

            generate_personalized_letters()

    elif file_input_choice == "Type Letter Content":
        show_placeholder_instructions()

        # Get letter content via text input
        letter_content = simpledialog.askstring(
            title="Mail Merge",
            prompt="Enter the body of the letter"
        )

        if not letter_content:
            show_exit_confirmation(process_letter_content)

        # Validate [name] placeholder exists
        name_place = letter_content.find("[name]")
        if name_place < 0:
            Messagebox.show_warning(
                title="Mail Merge",
                message="Warning: Placeholder '[name]' is not in the letter body"
            )
            is_placeholder_not_present = True
        else:
            is_placeholder_not_present = False

        if is_placeholder_not_present:
            # Re-prompt for file input method
            process_letter_content()

        generate_personalized_letters()


# =============================================================================
# INITIALIZATION & MAIN FLOW
# =============================================================================

# Create hidden main window (used as parent for dialogs)
app_window = ttk.Window(themename="darkly")
app_window.attributes('-alpha', 0)
app_window.iconbitmap("logo.ico")
center_window(app_window)

# Initialize prompts
name_input_method_prompt = "Do you want to enter names manually, or through a text file?"
recipient_count_prompt = "How many people do you want send a mail to?"
name_prompt = "Please enter the name:"

# Initialize variables
letter_content = ""
recipient_names = []
is_placeholder_not_present = False

# Ask user for name input method
input_method_choice = ""

name_collection_method()


# =============================================================================
# MANUAL NAME ENTRY FLOW
# =============================================================================
if input_method_choice == "Manual Insert":
    total_recipients = 0
    prompt_recipient_count()

    recipient_names = []
    collect_names_manually()

    # Show confirmation dialog with collected names
    names_confirmation_dialog = MessageDialog(
        title="Mail Merge",
        message=f"Names entered:\n{', '.join(recipient_names)}",
        buttons=["Continue", "Re-enter Names", "Exit"]
    )

    file_input_choice = confirm_names_and_proceed()
    process_letter_content()

# =============================================================================
# TEXT FILE NAME ENTRY FLOW
# =============================================================================
elif input_method_choice == "Text File":
    Messagebox.show_info(
        title="Mail Merge",
        message="Please make sure to separate the names with a comma, i.e. \", \"."
    )

    name_text_file_processing()

    process_letter_content()