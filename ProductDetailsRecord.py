class ProductDetailsRecord:
    def __init__(self, modelno, inventory_id=None, search_keywords=None, product_name=None, 
                 description=None, features=None, includes=None, specs=None, images=None, 
                 video=None, sds=None):
        self.modelno = modelno
        self.inventory_id = inventory_id
        self.search_keywords = search_keywords
        self.product_name = product_name
        self.description = description
        self.features = features
        self.includes = includes
        self.specs = specs
        self.images = images
        self.video = video
        self.sds = sds

    def insertRecord(self):
        values = [
            self.modelno, self.inventory_id, self.search_keywords, self.product_name,
            self.description, self.features, self.includes, self.specs, self.images,
            self.video, self.sds
        ]
        return [value if value not in ("", None) else None for value in values]

    def buildUpdateQuery(self):
        fields = {
            'search_keywords': self.search_keywords,
            'product_name': self.product_name,
            'description': self.description,
            'features': self.features,
            'includes': self.includes,
            'specs': self.specs,
            'images': self.images,
            'video': self.video,
            'sds': self.sds
        }

        # Filter out fields with None values and convert empty strings to None
        non_null_fields = {key: value if value != "" else None for key, value in fields.items() if value is not None}

        if not non_null_fields:
            return None, []

        # Build the SET clause
        set_clause = ", ".join([f"{key} = %s" for key in non_null_fields])
        parameters = list(non_null_fields.values())

        # Append the inventory_id to the parameters
        parameters.append(self.inventory_id)

        # Construct the final SQL query
        update_query = f"""
            UPDATE ProductDetails
            SET {set_clause}
            WHERE inventory_id = %s;
        """
        
        return update_query, parameters

    def __str__(self) -> str:
        return (f"Model Number: {self.modelno}, Inventory ID: {self.inventory_id}, "
                f"Search Keywords: {self.search_keywords}, Product Name: {self.product_name}, "
                f"Description: {self.description}, Features: {self.features}, "
                f"Includes: {self.includes}, Specs: {self.specs}, Images: {self.images}, "
                f"Video: {self.video}, SDS: {self.sds}")
