# xls2inv #

xls2inv is a "serverless" Python function and JavaScript front end that parses credit card order logs in Excel format and converts them to .inv format, a plain text file format used by Sierra ILS to batch upload serials invoices.

The Python script is designed to run on AWS Lambda. The JavaScript front end can be deployed from an AWS S3 bucket configured to host a static website.

### Requirements ###

#### Backend ####
* [openpyxl Python library](http://openpyxl.readthedocs.io)
* [boto3 Python library](https://boto3.readthedocs.io)
* [AWS S3](https://aws.amazon.com/s3/) bucket (for Excel/INV files)
* [AWS Lambda](https://aws.amazon.com/lambda/)
* A spreadsheet in .xlsx format, based on template

#### Frontend ####
* [AWS IAM](https://aws.amazon.com/iam/) identity pool/[AWS Cognito](https://aws.amazon.com/cognito/) role w/backend S3 bucket LIST/PUT/GET permissions
* [AWS SDK for JavaScript](https://aws.amazon.com/sdk-for-browser/)

### How do I deploy this application? ###

#### Backend ####
##### 1. S3 Bucket #####
Create a bucket in S3 for storing your Excel file uploads and INV output files.

In the Permissions tab of your bucket, click the "CORS configuration" button and paste the following code into the editor and click "Save":

```xml
<?xml version="1.0" encoding="UTF-8"?>
<CORSConfiguration xmlns="http://s3.amazonaws.com/doc/2006-03-01/">
    <CORSRule>
        <AllowedOrigin>*</AllowedOrigin>
        <AllowedMethod>POST</AllowedMethod>
        <AllowedMethod>GET</AllowedMethod>
        <AllowedMethod>PUT</AllowedMethod>
        <AllowedMethod>DELETE</AllowedMethod>
        <AllowedMethod>HEAD</AllowedMethod>
        <AllowedHeader>*</AllowedHeader>
    </CORSRule>
</CORSConfiguration>
```

##### 2.Create Lambda Package #####
Compress the xls2inv.py script into a .zip archive also containing the contents of your virtual environment's site-packages folder (including openpyxl) by running the following commands:
```bash
cd path/to/site-packages
zip -r9 ~/xls2inv.zip *
cd path/to/xls2inv.py
zip -g ~/xls2inv.zip xls2inv.py
```

##### 3. Create a Lambda function in AWS Console #####
From the Lambda dashboard, click on "Create a Lambda function."

On the "Select Blueprint" page, click on "Blank Function."

On the "Configure triggers" page, select S3 from the list of integrations.

In the "Bucket" dropdown, select the bucket you created in step one.

In the "Event type" dropdown, select "Object Created (All)."

In the "Prefix" textbox, enter 'input/' (without quotation marks). (The frontend will place uploaded sheets in the bucket's 'input' subfolder, so your trigger needs to watch that folder.)

In the "Suffix" textbox, enter '.xlsx' (without quotation marks) so the application will only run on Excel files.

Select the 'Enable trigger' checkbox and click the "Next" button.

On the "Configure function" page, enter whatever name and description you wish. (This is how your function will display in Lambda's function list.)

In the Runtime dropdown, select Python 2.7.

In the "Code entry type" dropdown, select "Upload a .ZIP file" and upload the xls2inv.zip package you created in step two.

In the "Handler" textbox, enter 'xls2inv.handler' (without the quotation marks). This is name of your Python script (xls2inv) and the function within that script that runs your code on an uploaded file (handler).

In the "Role" textbox select "Create a custom role." This will open AWS IAM in a new browser tab and allow you to automatically create a role with the necessary permissions to access uploaded files in your S3 bucket and place parsed invoice files in the bucket when the script completes.

Click the "Next" button.

At the bottom of the "Review" page, click the "Create function" button.

#### Frontend ####

##### 1. Cognito and IAM #####
The frontend requires an AWS Cognito identity pool associated with an AWS IAM role so unauthenticated AWS users can use the web interface to convert files.

Go to the Cognito console and click the "Manage Federated Identities button," then click the "Create a new identity pool" button.

Enter a name into the "Identity pool name" textbox, select the "Enable access to unauthenticated identities" checkbox, and click the "Create Pool" button.

Expand the "View Details" option on the next page. You should see the option to create a new IAM role for both authenticated and unauthenticated identities. Leave "Create a new IAM Role" selected in both dropdowns. You can use the default Role Names or give them new names. For each role, click the "View Policy Document" option and paste the following code into the editors (substituting your backend bucket's name for 'BUCKET_NAME'):
```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Effect": "Allow",
            "Action": [
                "s3:*"
            ],
            "Resource": [
                "arn:aws:s3:::BUCKET_NAME",
                "arn:aws:s3:::BUCKET_NAME/*"
            ]
        }
    ]
}
```

Click the "Allow" button after pasting the policy code into both editors.

On the "Sample code" page, make a note of the Identity Pool ID in red characters inside the "Get AWS Credentials" box. (You'll need this ID later.)

##### 2. Backend S3 Bucket Permissions #####
In order for user of the web frontend to view a list of invoice files, download an invoice file, and upload spreadsheets to be converted, your identity pool must be granted the proper permissions.

Go to the S3 bucket you created in the backend instructions and again go to the "Permissions" tab. This time, select "Bucket Policy" button. Below the editor window, click the "Policy generator" link.

On the "AWS Policy Generator" page, select "S3 Bucket Policy" in the policy type dropdown.

In the "Add Statements" section of the page, enter the following settings:

Effect: Allow
Principal: *
Actions: "GetObject" and "PutObject"
ARN: arn:aws:s3:::BUCKET_NAME (substituting your backend bucket's name for 'BUCKET_NAME')

Then click the "Add Statement" button.

Before generating your policy, go back to the "Add Statements" section and enter the following settings:

Effect: Allow
Principal: *
Actions: "ListBucket"
ARN: arn:aws:s3:::BUCKET_NAME (substituting your backend bucket's name for 'BUCKET_NAME')

Click the "Add Statement" button again. Then click the "Generate Policy" button.

Copy the generated Policy JSON Document and paste it into the Bucket Policy Editor on your S3 bucket's Permissions tab. Before saving, you need to make one edit to the generated policy. In the policy with the Actions of "s3:GetObject" and "S3:PutObject," change the "Resource" value from "arn:aws:s3:::BUCKET_NAME" to "arn:aws:s3:::BUCKET_NAME/\*" (substituting your backend bucket's name for BUCKET_NAME).

Click the "Save" button.

##### 3. Configure app.js #####
Before deploying the frontend, you need to configure app.js to use your Incognito identity pool and your backend S3 bucket.

Open app.js in a text/code editor and edit the following three lines of code replacing 'BUCKET_NAME' with your backend bucket's name, 'REGION' with your bucket's region, and 'IDENTITY_POOL_ID' with the ID of the Cognito identity pool you created above in step one:

```javascript
var bucketName = 'BUCKET_NAME';
var bucketRegion = 'REGION';
var IdentityPoolId = 'IDENTITY_POOL_ID';
```

##### 4. Create a frontend S3 bucket #####
To deploy your frontend, you will need to upload the files to an S3 Bucket and configure it to host a static website.

Create a new S3 bucket. (This needs to be a separate bucket from the one you created for your backend.) Then go to the bucket's "Properties" tab and click on "Static website hosting".

Select the "Use this bucket to host a website" option and enter "index.html" (without quotation marks) in the "Index document" textbox.

Make a note of the "Endpoint" URL. This is the web address of your frontend application.

Click the "Save" button.

##### 5. Upload the front end files #####
In the root level of the bucket you created in the last step, upload all three of the files in the 'frontend' directory of this repository:

* index.html
* app.js
* style.css

##### [Optional] 6. Upload the Excel template #####
If you wish, you can include a link to the Excel template included in this repository. This file is configured with the proper columns in the proper order to be converted by the Python script to INV format.

To make it available for users, upload it to the root of your frontend S3 bucket (or wherever you wish to host the file), then add a link to it somewhere in index.html or app.js.

### How do I convert a spreadsheet to .inv format ###

1. Navigate to the index.html address of your frontend S3 bucket.
1. Use the file uploader to upload a properly formatted XLSX file.
1. Click link to converted INV file to download formatted invoice.

### How should a spreadsheet be formatted? ###

* Spreadsheets must be in .xlsx format.
* Each row must contain a Sierra order record number.
* Spreadsheets must contain 500 or fewer rows, not including header row. (This limitation is not enforced by the application, but Sierra will not accept .inv files with more than 500 line items.)
* Spreadsheets must contain a header row. (The application ignores row 1 of all worksheets.)
* Spreadsheet data must be contained in a single worksheet titled 'Sheet1.'
* Data must follow template column order, numbered from left. (Application ignores header values.)
* In first row of last column, select procurement card user code for appropriate staff member.
* Refunds/rebates and other negative dollar values must be preceded by a negative sign.
* All monetary values must be in U.S. dollars.
* Only the first 29 characters of the "notes" column will be converted.

Data must appear in the following column order:
1. ORDER DATE
1. ORDER NUMBER
1. \# OF COPY
1. PRICE($)
1. S/H CHARGE &/OR SALES TAX (%)
1. TOTAL COST ($)
1. STAFF CODE (row 2 only; select field; max. 7 characters)

### Who do I contact for assistance? ###

* Developed by Tom Boone, Georgetown Law Library
