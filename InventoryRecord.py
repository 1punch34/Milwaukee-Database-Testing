class InventoryRecord:
    def __init__(self, MFG_Part_Number, GTIN=None, Status=None, Short_Description=None, 
                 Show_Online_Date=None, Available_to_Ship_Date=None, Brand=None, SubBrand=None, 
                 CountryCode=None, Warranty=None, Recalled=0, ReplacementModelNo=None, 
                 PreviousModelNo=None, PreviousUPC=None, OrderUOM=None, MinOrder=None, MulOrder=None, id=None):
        self.MFG_Part_Number = MFG_Part_Number
        self.GTIN = GTIN
        self.Status = Status
        self.Short_Description = Short_Description
        self.Show_Online_Date = Show_Online_Date
        self.Available_to_Ship_Date = Available_to_Ship_Date
        self.Brand = Brand
        self.SubBrand = SubBrand    
        self.CountryCode = CountryCode
        self.Warranty = Warranty
        self.Recalled = Recalled
        self.ReplacementModelNo = ReplacementModelNo
        self.PreviousModelNo = PreviousModelNo
        self.PreviousUPC = PreviousUPC
        self.OrderUOM = OrderUOM
        self.MinOrder = MinOrder
        self.MulOrder = MulOrder
        self.id = id  # This will be set after insertion if needed

    def insertRecord(self):
        return (
            self.MFG_Part_Number, self.GTIN, self.Status, self.Short_Description,
            self.Show_Online_Date, self.Available_to_Ship_Date, self.Brand, self.SubBrand,
            self.CountryCode, self.Warranty, self.Recalled, self.ReplacementModelNo,
            self.PreviousModelNo, self.PreviousUPC, self.OrderUOM, self.MinOrder, self.MulOrder
        )

    def buildUpdateQuery(self):
        # Define the dictionary of field names and values
        fields = {
            'GTIN': self.GTIN,
            'Status': self.Status,
            'Short_Description': self.Short_Description,
            'Show_Online_Date': self.Show_Online_Date,
            'Available_to_Ship_Date': self.Available_to_Ship_Date,
            'Brand': self.Brand,
            'SubBrand': self.SubBrand,
            'CountryCode': self.CountryCode,
            'Warranty': self.Warranty,
            'Recalled': self.Recalled,
            'ReplacementModelNo': self.ReplacementModelNo,
            'PreviousModelNo': self.PreviousModelNo,
            'PreviousUPC': self.PreviousUPC,
            'OrderUOM': self.OrderUOM,
            'MinOrder': self.MinOrder,
            'MulOrder': self.MulOrder
        }

        # Filter out fields with None values
        non_null_fields = {key: value for key, value in fields.items() if value is not None}

        if not non_null_fields:
            return None, []

        # Build the SET clause
        set_clause = ", ".join([f"{key} = %s" for key in non_null_fields])
        parameters = list(non_null_fields.values())

        # Append the MFG_Part_Number to the parameters
        parameters.append(self.MFG_Part_Number)

        # Construct the final SQL query
        update_query = f"""
            UPDATE inventory
            SET {set_clause}
            WHERE MFG_Part_Number = %s;
        """
        
        return update_query, parameters
    
    def __str__(self) -> str:
        return (f"Part Number: {self.MFG_Part_Number}, GTIN: {self.GTIN}, Status: {self.Status}, "
                f"Short Description: {self.Short_Description}, Show Online Date: {self.Show_Online_Date}, "
                f"Available to Ship Date: {self.Available_to_Ship_Date}, Brand: {self.Brand}, "
                f"SubBrand: {self.SubBrand}, Country Code: {self.CountryCode}, Warranty: {self.Warranty}, "
                f"Recalled: {self.Recalled}, Replacement Model No: {self.ReplacementModelNo}, "
                f"Previous Model No: {self.PreviousModelNo}, Previous UPC: {self.PreviousUPC}, "
                f"Order UOM: {self.OrderUOM}, Min Order: {self.MinOrder}, Mul Order: {self.MulOrder}, "
                f"ID: {self.id}")
