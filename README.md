# Google Drive Presentation Merger

A powerful Google Apps Script that automatically merges multiple PowerPoint (`.pptx`) or Google Slides files from a specified Google Drive folder into a single new Google Slides presentation.

This script is designed to handle large jobs by working around Google's 6-minute execution limit, using triggers to run in batches until the merge is complete.

## Features

* **Combine Multiple Presentations**: Merges dozens of `.pptx` and Google Slides files into one.
* **Multiple Sorting Options**: Merge files based on name, upload date, or modification date, in either ascending or descending order.
* **Robust Error Handling**: Uses a retry mechanism to handle API timeouts on large or complex slides.
* **Handles Large Jobs**: Automatically runs in 5-minute batches to avoid Google's 6-minute execution time limit.
* **Universal Naming Support**: The "sort by name" feature uses a natural sort algorithm that correctly sorts files like `Topic 2` and `Topic 10`.

## How to Use

1.  **Create the Script**:
    * Go to [script.google.com](https://script.google.com) and create a **New project**.
    * Copy the entire contents of the `Code.gs` file from this repository and paste it into the script editor, replacing any existing code.

2.  **Enable the Drive API**:
    * In the script editor, click the **+** icon next to **Services**.
    * Select **Drive API** and click **Add**.

3.  **Configure the Script**:
    * At the top of the script, find the `CONFIGURATION` section.
    * Replace `'PASTE_YOUR_FOLDER_ID_HERE'` with the ID of your Google Drive folder. You can find this in the folder's URL.

4.  **Run the Script**:
    * Save the project (💾 icon).
    * From the function dropdown at the top, select the function that matches your desired merge order (e.g., `mergeByName_Ascending`).
    * Click **Run**.
    * The first time, you will need to grant the script several permissions to access your Drive and Slides and to run automatically. This is normal.
    * The script will begin the merge process. For large jobs, it will run automatically in the background. You can check its progress on the **Executions** page in the Apps Script editor.

5.  **Stopping the Process**:
    * If you need to cancel a merge, run the `cancelMergeProcess` function.

## License

This project is open source and available under the [MIT License](LICENSE).