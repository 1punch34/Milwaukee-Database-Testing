import mysql.connector
import dateutil
import openpyxl.workbook
import LogModule
from QueryModule import QueryModule
from InventoryRecord import InventoryRecord
from ProductDetailsRecord import ProductDetailsRecord
from UOMRecord import UOMRecord
import MilwaukeeFileReader
from PriceRecord import PriceRecord
import MilwaukeePriceListReader

logger = LogModule.Logger(__name__).logger


def inventoryQuery(inventoryRecord : InventoryRecord, queryModule : QueryModule):
    exists = queryModule.inventoryRecordFetch(inventoryRecord.MFG_Part_Number)
    if exists: 
        queryModule.updateInventoryRecord(inventoryRecord)        # Update existing record
    else:
        queryModule.insertInventoryRecord(inventoryRecord)         # insert inventory record

def productInfoQuery(productDetailsRecord : ProductDetailsRecord, queryModule : QueryModule):
    #checks if the model number exists in prouduct details and inventory
    exists = queryModule.productDetailsRecordFetch(productDetailsRecord.modelno)
    inventoryID = queryModule.getInventoryID(productDetailsRecord.modelno)  
    if inventoryID:
        productDetailsRecord.inventory_id = inventoryID #get inventory ID
        if exists:
            queryModule.updateProductDetails(productDetailsRecord) # if exists in product deatils update record
        else:
            queryModule.insertProductDetails(productDetailsRecord) # else insert new record
    else:
        queryModule.insertInventoryRecord(InventoryRecord(productDetailsRecord.modelno)) # if not int inventory creates new inventory record
        inventoryID = queryModule.getInventoryID(productDetailsRecord.modelno)
        productDetailsRecord.inventory_id = inventoryID
        logger.warning(f"NEW INVENTORY RECORD CREATED : {productDetailsRecord.modelno}")
        queryModule.insertProductDetails(productDetailsRecord)
    
def UOMquery(uomRecord: UOMRecord, queryModule : QueryModule):

    inventoryID = queryModule.getInventoryID(uomRecord.mfg_part_no)
    exists = queryModule.uomRecordFetch(inventoryID, uomRecord.uom)
    if inventoryID:
        if exists:
            queryModule.updateUOMRecord(uomRecord)
        else:
            uomRecord.inventory_id = inventoryID
            queryModule.insertUOMRecord(uomRecord)
    else:
        queryModule.insertInventoryRecord(InventoryRecord(uomRecord.mfg_part_no))
        inventoryID = queryModule.getInventoryID(uomRecord.mfg_part_no)
        uomRecord.inventory_id = inventoryID
        logger.warning(f"NEW INVENTORY RECORD CREATED : {uomRecord.mfg_part_no}")
        queryModule.insertUOMRecord(uomRecord)

def priceRecordQuery(priceRecord: PriceRecord, queryModule : QueryModule):
    inventoryID = queryModule.getInventoryID(priceRecord.modelno)
    
    if inventoryID:
        priceRecord.inventory_id = inventoryID
        # Check if the record exists in the PriceTable
        exists = queryModule.priceRecordFetch(inventoryID)
        if exists:
            queryModule.updatePriceRecord(priceRecord)
        else:
            queryModule.insertPriceRecord(priceRecord)
    else:
        queryModule.insertInventoryRecord(InventoryRecord(priceRecord.modelno))
        inventoryID = queryModule.getInventoryID(priceRecord.modelno)
        priceRecord.inventory_id = inventoryID
        logger.warning(f"NEW INVENTORY RECORD CREATED : {priceRecord.modelno}")
        queryModule.insertPriceRecord(priceRecord)

def connect():
    try:
        conn = mysql.connector.connect(user='sanjid', password='sanjid1',
                                    host='192.168.0.10',
                                    database='productdata')
        return conn
    except (mysql.connector.Error, IOError) as err:
        logger.info("Failed to connect: %s", err)
        return None

def addFiletoDatabase(excelfile, fileReader):
  
    if fileReader == 'Milwaukee Product Information':
        print(fileReader)
        MilwaukeeFileReader.readFile(excelfile)
    elif fileReader == 'Milwaukee Price List':
        MilwaukeePriceListReader.readFile(excelfile)
    
    logger.info(LogModule.LogData.getLogData())






