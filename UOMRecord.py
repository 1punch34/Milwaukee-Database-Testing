class UOMRecord:
    def __init__(self, mfg_part_no, uom, upc, desc=None, quantity=None, weight=None, length =None, width=None, depth=None, height=None, inventory_id=None):
        self.mfg_part_no = mfg_part_no
        self.desc = desc
        self.quantity = quantity
        self.uom = uom
        self.upc = upc
        self.weight = weight
        self.length = length
        self.width = width
        self.depth = depth
        self.height = height
        self.inventory_id = inventory_id

    def insertRecord(self):
        return (self.mfg_part_no, self.desc, self.quantity, self.uom, self.upc, self.weight, self.length, self.width, self.depth, self.height, self.inventory_id)
    
    def updateQuery(self):
        fields = {
            'desc': self.desc,
            'quantity': self.quantity,
            'upc': self.upc,
            'weight': self.weight,
            'length' : self.length,
            'width': self.width,    
            'depth': self.depth,
            'height': self.height,
        }
        # Filter out None values
        non_null_fields = {k: v for k, v in fields.items() if v is not None}
        
        if not non_null_fields:
            return None, []

        set_clause = ", ".join([f"{key} = %s" for key in non_null_fields])
        update_query = f"UPDATE uoms SET {set_clause} WHERE inventory_id = %s AND uom = %s"

        values = list(non_null_fields.values())
        values.extend([self.inventory_id, self.uom])

        return update_query, values

    def __str__(self) -> str:
        return (f"Manufacturer Part Number: {self.mfg_part_no}, UOM: {self.uom}, UPC: {self.upc}, "
            f"Description: {self.desc}, Quantity: {self.quantity}, Weight: {self.weight}, "
            f"Width: {self.width}, Depth: {self.depth}, Height: {self.height}, "
            f"Inventory ID: {self.inventory_id}")    
