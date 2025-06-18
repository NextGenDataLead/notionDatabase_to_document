# Notion to Word & Google Docs Converter

This project provides a Python script to extract rich content from a Notion database and convert it into a Microsoft Word (`.docx`) document. Additionally, it can upload the generated Word document to Google Drive, converting it into a Google Doc.

## Features

*   **Notion Content Extraction**: Fetches pages from a specified Notion database.
*   **Rich Text Formatting**: Preserves bold, italic, strikethrough, underline, and basic code formatting.
*   **Image Handling**: Downloads and embeds images from Notion pages, with scaling to fit page width.
*   **List Support**: Converts bulleted and numbered lists, including nested lists.
*   **To-Do Items**: Represents Notion to-do items as checkboxes in the Word document.
*   **Customizable Output**: Allows specifying the output document name via CLI.
*   **Google Docs Integration**: Uploads the generated `.docx` file to Google Drive and converts it to a Google Doc.
*   **Filter History**: Remembers previous database IDs and filter configurations for quick reuse.
*   **Total Estimation Summary**: Calculates and prints the total estimated hours from "Estimation" properties of processed tickets.

## Prerequisites

Before running the script, ensure you have the following:

1.  **Python 3.9**: Installed on your system or with a conda environment.
2.  **Notion Integration Token**:
    *   Go to your Notion workspace settings -> Integrations -> Develop your own integrations.
    *   Click "+ New integration", give it a name (e.g., "Notion Doc Exporter"), and submit.
    *   Copy the "Internal Integration Token".
    *   Share your Notion database with this integration.
3.  **Google Cloud Project & Credentials**:
    *   Go to the [Google Cloud Console](https://console.cloud.google.com/).
    *   Create a new project or select an existing one.
    *   Enable the "Google Drive API" for your project.
    *   Go to "Credentials" -> "Create Credentials" -> "OAuth client ID".
    *   Choose "Desktop app" as the application type.
    *   Download the `client_secret.json` file and place it in the root directory of this project.
4.  **Notion Database ID**: The ID of the Notion database you want to extract content from. You can find this in the URL of your Notion database (e.g., `https://www.notion.so/your_workspace/DATABASE_ID?v=...`).

## Installation

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/your-username/Notion3.git
    cd Notion3
    ```
    (Note: Replace `https://github.com/your-username/Notion3.git` with the actual repository URL if different.)

2.  **Create a virtual environment (recommended)**:
    ```bash
    python3 -m venv venv
    source venv/bin/activate # On Windows: .\venv\Scripts\activate
    ```

3.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
    (You will need to create a `requirements.txt` file if it doesn't exist. See "Usage" for required packages.)

## Usage

### 1. Set up Environment Variables

Create a `.env` file in the project root and add your Notion API token:

```
NOTION_API_TOKEN="your_notion_integration_token_here"
```

Alternatively, you can pass the token directly via the `--token` CLI argument.

### 2. Install Python Dependencies

If `requirements.txt` does not exist, create it with the following content:

```
notion-client
python-docx
Pillow
python-dotenv
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
requests
```
Then run `pip install -r requirements.txt`.

### 3. Run the Script

You can run the script from your terminal.

**Basic Usage (will prompt for Database ID and Document Name):**

```bash
python3 notion_to_word.py
```

**With Arguments:**

```bash
python3 notion_to_word.py --database_id "YOUR_NOTION_DATABASE_ID" --document_name "MyProjectDocs"
```

**Available Arguments:**

*   `--token`: Your Notion API token. If not provided, the script will look for `NOTION_API_TOKEN` in `.env` or environment variables, or prompt you.
*   `--database_id`: The ID of the Notion database. If not provided, the script will offer previous IDs or prompt you.
*   `--document_name`: The base name for the output Word document (e.g., `MyProjectDocs.docx`) and the Google Doc (e.g., `MyProjectDocs_GoogleDoc`). If not provided, you will be prompted, with "NotionContent" as the default placeholder.
*   `--output_file`: (Advanced) The full path and filename for the output Word document. Defaults to `Output/[document_name]_[timestamp].docx`.
*   `--filter_history_file`: (Advanced) Path to the filter history JSON file.
*   `--db_history_file`: (Advanced) Path to the database ID history JSON file.

### Google Drive Authentication

The first time you run the script with Google Docs integration, a browser window will open asking you to authenticate with your Google account. Follow the prompts to grant access. A `token.json` file will be created to store your credentials for future runs.

## Project Structure

*   `notion_to_word.py`: The main script for extracting Notion content and generating the Word document.
*   `notion_to_gdoc.py`: Handles the Google Drive authentication and uploading/conversion of the Word document to Google Docs.
*   `client_secret.json`: Your Google API client secret file (downloaded from Google Cloud Console).
*   `token.json`: (Generated after first Google authentication) Stores your Google Drive API tokens.
*   `notion_filter_history.json`: (Generated) Stores your recent Notion filter configurations.
*   `notion_db_history.json`: (Generated) Stores your recent Notion database IDs.
*   `Output/`: Directory where generated Word documents are saved.
*   `Output/NotionContent_YYYYMMDD_HHMM.docx`: (Generated) Example of a default output Word document.

## Troubleshooting

*   **"Error fetching database info: object dict can't be used in 'await' expression"**: Ensure you have `notion-client` installed and that the script is using `AsyncClient` and `await` correctly. This has been addressed in the latest script version.
*   **"Permission denied" during Google Docs upload**: Check your Google Cloud project's credentials and ensure the Google Drive API is enabled and your OAuth client ID has the necessary permissions. You might need to delete `token.json` to re-authenticate.
*   **`client_secret.json` not found**: Make sure you have downloaded the `client_secret.json` file from Google Cloud Console and placed it in the project's root directory.
*   **Notion API Token issues**: Double-check your `NOTION_API_TOKEN` in the `.env` file or ensure you are providing it correctly via the `--token` argument. Also, ensure your Notion integration has access to the database you are trying to query.
