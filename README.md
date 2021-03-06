# xls2inv #

xls2inv is a "serverless" Python function and JavaScript front end that parses credit card order logs in Excel format and converts them to [.inv format](http://vendordocs.iii.com/#serials_elec_invoicing.html), a plain text file format used by Sierra ILS to batch upload serials invoices.

The Python script is designed to run on AWS Lambda. The JavaScript front end can be deployed from an AWS S3 bucket configured to host a static website.

![Screenshot of xls2inv frontend](screenshot.png)

### Requirements ###

#### Backend ####
* [openpyxl Python library](http://openpyxl.readthedocs.io)
* [boto3 Python library](https://boto3.readthedocs.io)
* [AWS S3](https://aws.amazon.com/s3/) bucket (for Excel/INV files)
* [AWS Lambda](https://aws.amazon.com/lambda/)

#### Frontend ####
* [AWS IAM](https://aws.amazon.com/iam/) identity pool/[AWS Cognito](https://aws.amazon.com/cognito/) role w/backend S3 bucket LIST/PUT/GET permissions
* [AWS SDK for JavaScript](https://aws.amazon.com/sdk-for-browser/)
* [AWS S3](https://aws.amazon.com/s3/) bucket (for frontend files)


### How do I deploy this application? ###

#### Backend ####
##### A. S3 Bucket #####

1. Create a bucket in S3 for storing your Excel file uploads and INV output files.

2. In the Permissions tab of your bucket, click the "CORS configuration" button and paste the following code into the editor and click "Save":  
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

##### B. Create Lambda Package #####

1. A Python package for deployment to AWS Lambda must include any add-on libraries utilized in the application. For xls2inv, that means boto3 (AWS's SDK library for Python) and openpyxl (the Excel parsing library). To include these, create a Python virtual environment on your local machine and install the necessary libraries within it:  
```bash
virtualenv -p python2.7 xls2inv_env
cd xls2inv_env
source bin/activate
pip install boto3
pip install openpyxl
```

2. With the libraries ready, you can compress the xls2inv.py script into a .zip archive along with the contents of your virtual environment's site-packages folder (which now includes boto3 and openpyxl):  
```bash
cd path/to/xls2inv_env/lib/python2.7/site-packages
zip -r9 ~/xls2inv.zip *
cd path/to/xls2inv.py
zip -g ~/xls2inv.zip xls2inv.py
```

##### C. Create a Lambda function in AWS Console #####
1. From the Lambda dashboard, click on "Create a Lambda function."

2. On the "Select Blueprint" page, click on "Blank Function."

3. On the "Configure triggers" page, select S3 from the list of integrations.

4. In the "Bucket" dropdown, select the bucket you created in step one.

5. In the "Event type" dropdown, select "Object Created (All)."

6. In the "Prefix" textbox, enter 'input/' (without quotation marks). (The frontend will place uploaded sheets in the bucket's 'input' subfolder, so your trigger needs to watch that folder.)

7. In the "Suffix" textbox, enter '.xlsx' (without quotation marks) so the application will only run on Excel files.

8. Select the 'Enable trigger' checkbox and click the "Next" button.

9. On the "Configure function" page, enter whatever name and description you wish. (This is how your function will display in Lambda's function list.)

10. In the Runtime dropdown, select Python 2.7.

11. In the "Code entry type" dropdown, select "Upload a .ZIP file" and upload the xls2inv.zip package you created in step two.

12. In the "Handler" textbox, enter 'xls2inv.handler' (without the quotation marks). This is name of your Python script (xls2inv) and the function within that script that runs your code on an uploaded file (handler).

13. In the "Role" textbox select "Create a custom role." This will open AWS IAM in a new browser tab and allow you to automatically create a role with the necessary permissions to access uploaded files in your S3 bucket and place parsed invoice files in the bucket when the script completes.

14. Click the "Next" button.

15. At the bottom of the "Review" page, click the "Create function" button.

#### Frontend ####

##### A. Cognito and IAM #####

The frontend requires an AWS Cognito identity pool associated with an AWS IAM role so unauthenticated AWS users can use the web interface to convert files.

1. Go to the Cognito console and click the "Manage Federated Identities button," then click the "Create a new identity pool" button.

2. Enter a name into the "Identity pool name" textbox, select the "Enable access to unauthenticated identities" checkbox, and click the "Create Pool" button.

3. Expand the "View Details" option on the next page. You should see the option to create a new IAM role for both authenticated and unauthenticated identities. Leave "Create a new IAM Role" selected in both dropdowns. You can use the default Role Names or give them new names. For each role, click the "View Policy Document" option and paste the following code into the editors (substituting your backend bucket's name for 'BUCKET_NAME'):  
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

4. Click the "Allow" button after pasting the policy code into both editors.

5. On the "Sample code" page, make a note of the Identity Pool ID in red characters inside the "Get AWS Credentials" box. (You'll need this ID later.)

##### B. Backend S3 Bucket Permissions #####

In order for user of the web frontend to view a list of invoice files, download an invoice file, and upload spreadsheets to be converted, your identity pool must be granted the proper permissions.

1. Go to the S3 bucket you created in the backend instructions and again go to the "Permissions" tab. This time, select "Bucket Policy" button. Below the editor window, click the "Policy generator" link.

2. On the "AWS Policy Generator" page, select "S3 Bucket Policy" in the policy type dropdown.

3. In the "Add Statements" section of the page, enter the following settings:  
Effect: Allow  
Principal: \*  
Actions: "GetObject" and "PutObject"  
ARN: arn:aws:s3:::BUCKET_NAME/\* (substituting your backend bucket's name for 'BUCKET_NAME')

Then click the "Add Statement" button.

4. Before generating your policy, go back to the "Add Statements" section and enter the following settings:  
Effect: Allow  
Principal: \*  
Actions: "ListBucket"  
ARN: arn:aws:s3:::BUCKET_NAME (substituting your backend bucket's name for 'BUCKET_NAME')

5. Click the "Add Statement" button again. Then click the "Generate Policy" button.

6. Copy the generated Policy JSON Document and paste it into the Bucket Policy Editor on your S3 bucket's Permissions tab, and click the "Save" button.

##### C. Configure app.js #####

Before deploying the frontend, you need to configure app.js to use your Incognito identity pool and your backend S3 bucket.

1. Open app.js in a text/code editor and edit the following three lines of code replacing 'BUCKET_NAME' with your backend bucket's name, 'REGION' with your bucket's region, and 'IDENTITY_POOL_ID' with the ID of the Cognito identity pool you created above in step one:  
```javascript
var bucketName = 'BUCKET_NAME';
var bucketRegion = 'REGION';
var IdentityPoolId = 'IDENTITY_POOL_ID';
```

##### D. Create a frontend S3 bucket #####

To deploy your frontend, you will need to upload the files to an S3 Bucket and configure it to host a static website.

1. Create a new S3 bucket. (This needs to be a separate bucket from the one you created for your backend.) Then go to the bucket's "Properties" tab and click on "Static website hosting".

2. Select the "Use this bucket to host a website" option and enter "index.html" (without quotation marks) in the "Index document" textbox.

3. Make a note of the "Endpoint" URL. This is the web address of your frontend application.

4. lick the "Save" button.

##### E. Upload the front end files #####

1. In the root level of the bucket you created in the last step, upload all three of the files in the 'frontend' directory of this repository:  
* index.html
* app.js
* style.css

##### [Optional] F. Upload the Excel template #####
If you wish, you can include a link to the Excel template included in this repository. This file is configured with the proper columns in the proper order to be converted by the Python script to INV format.

To make it available for users, upload it to the root of your frontend S3 bucket (or wherever you wish to host the file), then add a link to it somewhere in index.html or app.js.

### How do I convert a spreadsheet to .inv format ###

1. Navigate to the index.html address of your frontend S3 bucket.
2. Use the file uploader to upload a properly formatted XLSX file.
3. Click link to converted INV file to download formatted invoice.

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
* Only the first 29 characters of the "NOTE" column will be used.

Data must appear in the following column order:
1. ORDER DATE
2. ORDER NUMBER
3. \# OF COPY
4. PRICE($)
5. S/H CHARGE &/OR SALES TAX (%)
6. TOTAL COST ($)
7. NOTE
8. STAFF CODE (row 2 only; first 7 characters used to generate header invoice ID)

### Who do I contact for assistance? ###

* Developed by Tom Boone, Georgetown Law Library
