import dotenv from 'dotenv';
dotenv.config();
import fs from 'node:fs';
import xlsx from 'xlsx';
xlsx.set_fs(fs)
import mapTransform, { alt, transform, transformers } from 'map-transform'
import { createApi } from './api.js';
const SIRION_CLIENTID = process.env.SIRION_CLIENTID
const SIRION_CLIENTSECRET = process.env.SIRION_CLIENTSECRET
const SIRION_LOGINID = process.env.SIRION_LOGINID // optional, if empty then permissions defined on OAuth credential are applied and activty is logged as system admin
const SIRION_URL = process.env.SIRION_URL
const api = createApi(SIRION_URL);
const FILES_DIR = 'files'
const CLI_BATCH_SIZE = 1000
const API_VER = 'v2'
const CLT_ID_OVERRIDE = '' // if populated then clt id is not required in excel filename

// transform excel cell values into sirion cli field values
// fixed values
const unitType = () => (d) => { return API_VER === 'v2' ? { id: 2098 } : { "s_uuid": "61da74d4-5f79-4f56-b075-0294ffc5672f" } }
const pricingType = () => (d) => { return API_VER === 'v2' ? { id: 1001 } : { "s_externalId": "Fixed" } }
const country = () => (d) => 'United States'
const currencyV2 = [{ n: "EUR", i: 2 }, { n: "INR", i: 8 }, { n: "TRY", i: 48 }, { n: "AED", i: 49 }, { n: "GBP", i: 4 }, { n: "USD", i: 1 }]
const currency = () => (d) => { return API_VER === 'v2' ? { id: currencyV2.find(c => c.n === d) } : { externalId: d } }
// transforms excel date number to dd.MM.yyyy
const excelToDate = () => (n) => { return new Date((n - 25569) * 86400 * 1000).toLocaleDateString('de-DE') };
// update field mappings with sirion clt id, field names and excel column names
const fieldMappings = []
fieldMappings['CLT02952'] = {
    $iterate: true,
    //'CLT field name': 'Excel column name'
    'Product': 'SKU',
    'Product Description': 'Product Name',
    'Quantity': 'Qty',
    'Price': 'Price',
    'Unit Type': alt(null, transform(unitType)),
    'Pricing Type': alt(null, transform(pricingType)),
    'Start Date': ['Effective Date', transform(excelToDate)],
    'End Date': ['Expiration Date', transform(excelToDate)]
}
fieldMappings['CLT03063'] = {
    $iterate: true,
    //'CLT field name': 'Excel column name'
    'Number': 'SKU',
    'Product': 'Product Name',
    'Quantity': 'Qty',
    'Price': 'Price',
    'Country': alt(null, transform(country)),
}

const getExcelData = async (filename) => {
    const workbook = xlsx.readFile(filename);
    // first worksheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);
    return data
}

const mapFields = async (data, cltId) => {
    if (!fieldMappings[cltId]) {
        throw new Error(`No fieldMappings found for ${cltId}`)
    }
    const mapper = mapTransform(fieldMappings[cltId])
    const target = await mapper(data)
    // cli require unique rowNum
    const result = target.map((row, index) => ({
        ...row,
        rowNum: index + 1,
    }));
    return result
}

const createContractLineItems = async (token, entityName, entityId, cltId, cliData) => {
    if (!Array.isArray(cliData) || !cliData.length > 0) {
        throw new Error('cliData must be an array')
    }
    // batch 1000 cli
    const responses = []
    for (let index = 0; index < cliData.length; index += CLI_BATCH_SIZE) {
        const batch = cliData.slice(index, index + CLI_BATCH_SIZE)
        if (API_VER === 'v2') {
            const response = await api.createContractLineItems(token, entityName, entityId, cltId, {
                data: batch
            })
            responses.push(response)
        } else {
            const cltUuid = await api.getContractLineItemTemplateUuid(token, entityName, cltId)
            const entityUuid = await api.getEntityUuid(token, entityName, entityId)
            const response = await api.createContractLineItemsV3(token, entityName, entityUuid, cltUuid, {
                data: batch
            })
            responses.push(response)
        }
    }
    return responses
}

(async () => {
    // sirion api authentication
    const token = await api.auth.clientToken(SIRION_CLIENTID, SIRION_CLIENTSECRET, SIRION_LOGINID)
    // iterate through excel files
    const excelFiles = fs.readdirSync(FILES_DIR)
        .filter((filename) => /\.(xls|xlsx)$/i.test(filename))
    for (const excelFile of excelFiles) {
        const entityId = excelFile.split('.')[0]
        const entityName = entityId.toUpperCase().startsWith('CO') ? 'contracts' : 'contract-draft-requests'
        const cltId = !CLT_ID_OVERRIDE ? excelFile.split('.')[1] : CLT_ID_OVERRIDE
        const excelData = await getExcelData(`${FILES_DIR}/${excelFile}`)
        const cliData = await mapFields(excelData, cltId)
        const responses = await createContractLineItems(token, entityName, entityId, cltId, cliData)
        console.log(excelFile)
        console.log(JSON.stringify(responses, null, 2).slice(0, 1000))
    }
})()
