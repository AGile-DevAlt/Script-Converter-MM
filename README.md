# Excel Sheet Converter

## Run locally
1. Install Python 3.10 or newer.
2. Open Terminal/Command Prompt in this folder.
3. Install dependencies:
   `python -m pip install -r requirements.txt`
4. Start the app:
   `streamlit run app.py`
5. Your browser will open the converter. Upload an `.xlsx` or `.xls` file, check the suggested column mappings, then click **Convert Excel** and download the result.

The converter uses column names to suggest mappings, but you should review them because different schools may label fields differently. Missing source fields are left blank. Your original file is not edited.
