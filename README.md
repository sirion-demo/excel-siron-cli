# Excel to Sirion Contract Line Items
Read Excel files and create Sirion contract line items with field name mapping, field value transformation and API rate limiting
## Usage 
```
npm install
npm start
```
## Tenant variables
Copy .env_example and rename to .env then set the following environment variables
```
SIRION_URL=https://<tenant>.sirioncloud.com
SIRION_CLIENTID=<clientId>
SIRION_CLIENTSECRET=<clientSecret>
```
## Excel files
Copy Excel files into the 'files' folder with these naming conventions    
If CLT_ID_OVERRIDE is not populated, then the CLT ID **is** required in the Excel filename, enabling multiple CLTs and both CO/CDR
```
const CLT_ID_OVERRIDE = ''
Excel filename convention: <Entity ID>.<CLT ID>.xlsx 
Examples: CO02027.CLT02952.xlsx, CDR06756.CLT03063.xlsx
 ```
If CLT_ID_OVERRIDE is populated, then the CLT ID is **not** required in the Excel filename and only this CLT is used
```
const CLT_ID_OVERRIDE = 'CLT02952'
Excel filename convention: <Entity ID>.xlsx 
Examples: CO02027.xlsx, CO02028.xlsx
```

## Update field transformations and mappings (src/index.mjs) 
Documentation https://www.npmjs.com/package/map-transform  
Modify to transform Excel cell values into Sirion field values or fixed values
```
// fixed values
const unitType = () => (d) => { return {id: 2098} }
const pricingType = () => (d) => { return {id: 1001} }
// transform Excel date number to dd.MM.yyyy
const excelToDate = () => (n) => { return new Date((n - 25569) * 86400 * 1000).toLocaleDateString('de-DE')};
```
Update field mappings with Sirion contract line item template ID, field names and Excel column names
```
const fieldMapping['<CLT ID>'] = {
    $iterate: true,
    //'CLT field name': 'Excel column name'
    'Product': 'SKU',
    'Product Description': 'Product Name',
    'Quantity': 'Qty',
    'Price': 'Price',
    'Unit Type':  alt(null, transform(unitType)), // fixed value
    'Pricing Type': alt(null, transform(pricingType)), // fixed value
    'Start Date':['Effective Date', transform(excelToDate)], // transforms excel date number to dd.MM.yyyy
    'End Date': ['Expiration Date', transform(excelToDate)] // transforms excel date number to dd.MM.yyyy
}
```
## XLSX Package
Documentation https://docs.sheetjs.com/docs/  
The xlxs package on npmjs.com contains security issues, this version doesn't have them
```
npm i --save https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz 
```