class PriceRecord:
    def __init__(self, modelno, description=None, status=None, order_quantity=None, uom=None, list_price=None, inventory_id=None):
        self.modelno = modelno
        self.description = description
        self.status = status
        self.order_quantity = order_quantity
        self.uom = uom
        self.list_price = list_price
        self.inventory_id = inventory_id

    def insertRecord(self):
        values = (
            self.modelno,
            self.description,
            self.status,
            self.order_quantity,
            self.uom,
            self.list_price,
            self.inventory_id
        )
        return [value if value not in ("", None) else None for value in values]
    
    def buildUpdateQuery(self):
        fields = {
            'description': self.description,
            'status': self.status,
            'order_quantity': self.order_quantity,
            'uom': self.uom,
            'list_price': self.list_price,
            'inventory_id': self.inventory_id
        }

        # Filter out None values
        non_null_fields = {key: (value if value != '' else None) for key, value in fields.items() if value not in (None, '')}

        if not non_null_fields:
            return None, []

        # Build the SET clause
        set_clause = ", ".join([f"{key} = %s" for key in non_null_fields])
        parameters = list(non_null_fields.values())

        # Append the modelno to the parameters
        parameters.append(self.modelno)

        # Construct the final SQL query
        update_query = f"""
            UPDATE PriceTable
            SET {set_clause}
            WHERE modelno = %s;
        """
        
        return update_query, parameters
    
    def __str__(self):
        return (f"PriceRecord(modelno={self.modelno}, description={self.description}, status={self.status}, "
                f"order_qty={self.order_quantity}, uom={self.uom}, list_price={self.list_price}, "
                f"inventory_id={self.inventory_id})")