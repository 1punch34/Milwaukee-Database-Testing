import dateutil.parser
import openpyxl
import mysql.connector
import time
import dateutil
import json
import openpyxl.workbook
import LogModule
from QueryModule import QueryModule
from InventoryRecord import InventoryRecord
from ProductDetailsRecord import ProductDetailsRecord
from UOMRecord import UOMRecord

logger = LogModule.Logger(__name__).logger


def inventoryQuery(row, header_indices, conn : mysql.connector.MySQLConnection):
    queryModule = QueryModule(conn)
    try:
        MFG_Part_Number = row[header_indices['MFG Part # (OEM)']]
        GTINCode = row[header_indices['GTIN']]
        status_index = header_indices.get('Status', None)
        Status = row[status_index] if status_index is not None and row[status_index] else None
        Short_Description = row[header_indices['Short Description']]
    except Exception as err:
        logger.error("HEADER ERROR in Inventory query")
    try:
        Show_Online_Date = dateutil.parser.parse(row[header_indices['Show Online Date']]).strftime('%Y-%m-%d')
        Available_to_Ship_Date = dateutil.parser.parse(row[header_indices['Available to Ship Date']]).strftime('%Y-%m-%d')
    except Exception as err:
        Show_Online_Date = row[header_indices['Show Online Date']]
        Available_to_Ship_Date = row[header_indices['Available to Ship Date']]
        logger.error(f"Error parsing Dates{Show_Online_Date, Available_to_Ship_Date} for row with with MFG_Part_Num: {MFG_Part_Number}: {err}")

    inventoryRecord = InventoryRecord(
            MFG_Part_Number=MFG_Part_Number, GTIN=GTINCode, Status=Status, Short_Description=Short_Description, 
            Show_Online_Date=Show_Online_Date, Available_to_Ship_Date=Available_to_Ship_Date,
        )
    
    exists = queryModule.inventoryRecordFetch(MFG_Part_Number)
    if exists:
        # Update existing record 
        queryModule.updateInventoryRecord(inventoryRecord)
    else:
        # insert inventory record
        queryModule.insertInventoryRecord(inventoryRecord)

def productInfoQuery(row, header_indices, conn : mysql.connector.MySQLConnection):
    queryModule = QueryModule(conn)
    try:
        modelno = row[header_indices['MFG Part # (OEM)']]
        search_keywords = row[header_indices['Search Keywords']]
        product_name = row[header_indices['Product Name']]
        description = row[header_indices['Marketing Copy']]
        features_list = [
            row[header_indices[f'Feature - Benefit Bullet {i}']] for i in range(1, 22)  if header_indices.get(f'Feature - Benefit Bullet {i}') and row[header_indices[f'Feature - Benefit Bullet {i}']]
        ]
        features = json.dumps(features_list)
        
        # Create the includes JSON object, filtering out any null values or empty strings
        includes_list = [row[header_indices['Package Contents']] if header_indices.get('Package Contents') and row[header_indices['Package Contents']] else '']
        includes = json.dumps(includes_list)
    except Exception as err:
        logger.error("HEADER ERROR in productInfo query")
    
    exists = queryModule.productDetailsRecordFetch(modelno)
    inventoryID = queryModule.getInventoryID(modelno)
    
    if inventoryID:
        productDetailsRecord = ProductDetailsRecord(modelno=modelno, inventory_id=inventoryID, search_keywords=search_keywords, product_name=product_name, 
                                               description=description, features=features, includes=includes)
        if exists:
            queryModule.updateProductDetails(productDetailsRecord)
        else:
            queryModule.insertProductDetails(productDetailsRecord)
    else:
        logger.error(f"No matching inventory found for modelno: {modelno}")

def digitalAssetsQuery(row, digitalassetsheader, conn : mysql.connector.MySQLConnection ):
    queryModule = QueryModule(conn)
    try:
        modelno = row[digitalassetsheader['MFG Part # (OEM)']]
        video = row[digitalassetsheader['Product Review Video']] 
        safetyDataSheet = row[digitalassetsheader['Safety Data Sheet (SDS) - PDF']]
        
        
        # Find index for the first hero image
        imageStart = digitalassetsheader.get('Main Product Image')
        if imageStart is None:
            raise KeyError('Main Product Image header is missing')

        # Find the index for the last Detailed Product View image
        detailed_views_indices = [idx for header, idx in digitalassetsheader.items() if header.startswith('Detailed Product View')]
        if not detailed_views_indices:
            raise KeyError('No Detailed Product View headers found')
        imagesEnd = max(detailed_views_indices)
    except Exception as err:
        logger.error("HEADER ERROR in Inventory query")
    # Extract all images between the first hero image and the last detailed product view
    imagesList = row[imageStart:imagesEnd + 1]
    imagesList= [img for img in imagesList if img]

    # Construct the JSON object
    images = json.dumps(imagesList)


    exists = queryModule.productDetailsRecordFetch(modelno)
    productDetailsRecord = ProductDetailsRecord(modelno=modelno, video=video, sds=safetyDataSheet, images=images
    )
    if exists:
        logger.info(f"UPDATING Digital Assets for model no {modelno}")
        queryModule.updateProductDetails(productDetailsRecord)
    else:
        try:
            inventoryID = queryModule.getInventoryID(modelno)
            if inventoryID:
                productDetailsRecord.inventory_id = inventoryID
                logger.warning(f"Inserting New Product Details Record from Digiatl Assets Sheet : {modelno}")
                queryModule.insertProductDetails(productDetailsRecord)
            else:
                logger.warning(f"Inserting New Inventory Record from Digiatl Assets Sheet : {modelno}")
                queryModule.insertInventoryRecord(InventoryRecord(modelno))
                inventoryID = queryModule.getInventoryID(modelno)
                productDetailsRecord.inventory_id = inventoryID
                logger.warning(f"Inserting New Product Details Record from Digiatl Assets Sheet : {modelno}")
                queryModule.insertProductDetails(productDetailsRecord)
        
        except mysql.connector.Error as err:
            logger.error(f"Error inserting into Inventory for modelno: {modelno}: {err}")

def UOMquery(mfg_part_no, desc, quantity, uom, upc, weight, width, depth, height, conn : mysql.connector.MySQLConnection):
    queryModule = QueryModule(conn)
    uomRecord = UOMRecord(mfg_part_no=mfg_part_no, uom=uom, upc=upc, desc=desc, quantity=quantity,
            weight=weight, width=width, depth=depth, height=height, inventory_id=inventoryID
        )
    
    inventoryID = queryModule.getInventoryID(mfg_part_no)
    if inventoryID:
        exists = queryModule.uomRecordFetch(inventoryID, uomRecord.uom)
        if exists:
            queryModule.updateUOMRecord(uomRecord)
        else:
            queryModule.insertUOMRecord(uomRecord)
    else:
        logger.error(f"No matching inventory found for mfg_part_no: {mfg_part_no}")

def addUOMs(row, header_indices, conn : mysql.connector.MySQLConnection):
    mfg_part_no = row[header_indices['MFG Part # (OEM)']] 
    desc = row[header_indices['Short Description']]

    Eachupc = row[header_indices['UPC']]
    package_height = row[header_indices['Package Height (In.)']]
    package_width = row[header_indices['Package Width (In.)']]
    package_depth = row[header_indices['Package Depth (In.)']]
    package_weight = row[header_indices['Package Weight (Lb.)']]

    UOMquery(mfg_part_no, desc, 1, 'EA', Eachupc, package_weight, package_width, package_depth, package_height, conn)

    # Inner pack details
    inner_pack_upc = row[header_indices['Inner Pack GTIN']]
    inner_pack_quantity = row[header_indices['Inner Pack Quantity']]
    inner_pack_height = row[header_indices['Inner Pack Height (In.)']]
    inner_pack_width = row[header_indices['Inner Pack Width (In.)']]
    inner_pack_depth = row[header_indices['Inner Pack Depth (In.)']]
    inner_pack_weight = row[header_indices['Inner Pack Weight (Lb.)']]
    
    if inner_pack_quantity:
        UOMquery(mfg_part_no, desc, inner_pack_quantity, f'IP{inner_pack_quantity}', inner_pack_upc, inner_pack_weight, inner_pack_width, inner_pack_depth, inner_pack_height, conn)

    # Case details
    case_upc = row[header_indices['Case GTIN']]
    case_quantity = row[header_indices['Case Quantity']]
    case_height = row[header_indices['Case Height (In.)']]
    case_width = row[header_indices['Case Width (In.)']]
    case_depth = row[header_indices['Case Depth (In.)']]
    case_weight = row[header_indices['Case Weight (Lb.)']]
    if case_quantity:
        UOMquery(mfg_part_no, desc, case_quantity, f'CASE{case_quantity}', case_upc, case_weight, case_width, case_depth, case_height, conn)

def addSpecs(row, header_indices, conn : mysql.connector.MySQLConnection):

    queryModule = QueryModule(conn)
    MFG_Part_Number = row[header_indices['MFG Part # (OEM)']]
    specDict = {}
    for header_name, index in header_indices.items():
        # Add the header name and value to the dictionary
        if row[index] is None:
            continue
        specDict[header_name] = row[index]

    specDict.pop("GTIN", None)
    specs = json.dumps(specDict)

    productDetails = ProductDetailsRecord(modelno= MFG_Part_Number, specs=specs)
    
    exists = queryModule.productDetailsRecordFetch
    if exists:
        logger.info(f"UPDATING SPEC for model No : {MFG_Part_Number}")
        queryModule.updateProductDetails(productDetails)
    else:
        try:
            inventoryID = queryModule.getInventoryID(MFG_Part_Number)
            if inventoryID:
                productDetails.inventory_id = inventoryID
                logger.warning(f"Inserting New productrecord from SPECSHEET for model No : {MFG_Part_Number}")
                queryModule.insertProductDetails(productDetails)
            else:
                logger.warning(f"NEW Invetnory RECORD CREATED from SPECSHEET: {MFG_Part_Number}")
                queryModule.insertInventoryRecord(InventoryRecord(MFG_Part_Number))
                inventoryID = queryModule.getInventoryID(MFG_Part_Number)
                productDetails.inventory_id = inventoryID
                queryModule.insertProductDetails(productDetails)
                logger.warning(f"INSERTED NEW Product details RECORD fromm SPECSHEET: {MFG_Part_Number}")
        
        except mysql.connector.Error as err:    
            logger.error(f"Error inserting NEW RECORD from SPECSHEET for modelno: {MFG_Part_Number}: {err}")
        
        # Insert into Inventory


def connect():
    try:
        conn = mysql.connector.connect(user='sanjid', password='sanjid1',
                                    host='192.168.0.10',
                                    database='productdata')
        return conn
    except (mysql.connector.Error, IOError) as err:
        logger.info("Failed to connect: %s", err)
        return None

def addFiletoDatabase(excelfile):

    workbook = openpyxl.open(excelfile)
    logger.info("FILE OPENED")
    conn = connect()
    if conn is not None and conn.is_connected():
        cursor = conn.cursor()
        cursor.execute("USE productdata")

        productInformationSheet = workbook['Product Information']
        headers = [cell.value for cell in productInformationSheet[1]]
        # Create a dictionary to map header names to column indices
        header_indices = {header: idx for idx, header in enumerate(headers)} 
        for row in productInformationSheet.iter_rows(min_row=2, max_row=productInformationSheet.max_row, values_only=True):
            MFG_Part_Number = row[header_indices['MFG Part # (OEM)']]
            if MFG_Part_Number is None:
                continue
            try:
                inventoryQuery(row , header_indices, conn)
                productInfoQuery(row , header_indices, conn)
                # addUOMs(row,header_indices,cursor,conn)
            except Exception as err:
                logger.error(f"Funciton Error @ProductInfoSheet: {err}")
            print("============================000000000000=============================  \n\n\n")


        digitalAssetsSheet = workbook['Digital Assets']
        digitalAssetsHeaders = [cell.value for cell in digitalAssetsSheet[1]]
        digitalAssetsHeaders_indices = {digitalAssetsHeader: idx for idx, digitalAssetsHeader in enumerate(digitalAssetsHeaders)}

        for row in digitalAssetsSheet.iter_rows(min_row=2, max_row=digitalAssetsSheet.max_row, values_only= True):
            MFG_Part_Number = row[header_indices['MFG Part # (OEM)']]
            if MFG_Part_Number is None:
                continue            
            try:
                digitalAssetsQuery(row,digitalAssetsHeaders_indices, conn)
            except Exception as err:
                logger.error(f"Funciton Error @DigitalAssetsSheet: {err}")
            print("============================000000000000=============================  \n\n\n")


        excluded_sheets = {"Product Information", "FR Product Information", "Digital Assets", "Digital Assets FR"}
        for sheet_name in workbook.sheetnames:
            if sheet_name in excluded_sheets:
                continue
            specSheet = workbook[sheet_name]
            specheaders = [cell.value for cell in specSheet[1]]
            # Create a dictionary to map header names to column indices
            specheader_indices = {header: idx for idx, header in enumerate(specheaders)} 

            for row in specSheet.iter_rows(min_row=2, max_row=specSheet.max_row, values_only= True):
                MFG_Part_Number = row[specheader_indices['MFG Part # (OEM)']]
                if MFG_Part_Number is None:
                    continue            
                try:
                    addSpecs(row, specheader_indices,conn)
                except Exception as err:
                    logger.error(f"Funciton Error @SpecSheet:{sheet_name} : {err}")
                print("============================000000000000=============================  \n\n\n")
        
        cursor.close()
        conn.close()




