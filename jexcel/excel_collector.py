import pandas as pd
import jexcel.header as hd
from jexcel.header import Header
from jexcel.header import HeaderType
from jexcel.header import PrimaryKey

class ExcelCollector:

    def __init__(self, df, headers, header_root=None):
        self.df = df
        self.headers = headers
        self.data_all = []
        self.reset_primary_history()
        self.header_root = header_root

        if header_root is None:
            if len(headers) > 0:
                self.header_root = headers[0].get_root()
        

    def reset_primary_history(self):
        self.primary_history = []
        for header in self.headers:
            if header.primary_key == header:
                self.primary_history.append(PrimaryKey(header))


    def update_primary_history(self, header, value):
        for primary_info in self.primary_history:
            if primary_info.header == header:
                primary_info.identy = value
                break
        for primary_info in self.primary_history:
            if primary_info.header.level > header.level:
                primary_info.identity = None

    
    def get_primary_history_value(self, header):
        for primary_info in self.primary_history:
            if primary_info.header == header.primary_key:
                return primary_info.identity
        return None


    def find_primary_item(self, primary_id, data_list, out_idx = -1):
        idx = -1
        for item in data_list:
            idx += 1
            key_name = primary_id.header.name
            if any(key_name in d for d in item) and item[key_name] == primary_id.identity:
                out_idx = idx
                return item
        out_idx = -1
        return None
    

    @staticmethod
    def parse_cell_value(cell_value):

        if cell_value is None:
            return None
        
        try:
            return int(cell_value)
        except ValueError:
            pass
        
        try:
            return float(cell_value)
        except ValueError:
            pass

        if isinstance(cell_value, str):
            if cell_value.lower() == "true":
                return True
            elif cell_value.lower() == "false":
                return False

        return cell_value
    

    @staticmethod
    def is_empty(data):

        if not data or data is None:
            return True
        
        if isinstance(data, dict):
            for key in data:
                if not ExcelCollector.is_empty(data[key]):
                    return False
            return True
        
        if isinstance(data, list):
            for item in data:
                if not ExcelCollector.is_empty(item):
                    return False
            return True
        
        return False


    def parse(self):

        if not self.headers or len(self.headers) == 0 or self.df.empty:
            return []

        column_primary = self.headers[0].name
        row = self.df.iloc[0]

        primary_ids = [PrimaryKey(self.header_root)]
        data = {}

        self.fill_data_row(primary_ids, data, row)
        self.data_all.append(data)

        for _, row in self.df.iloc[1:].iterrows():
            if not pd.isna(row[column_primary]):
                primary_ids = [PrimaryKey(self.header_root)]
                data = {}
                self.data_all.append(data)

            self.fill_data_row(primary_ids, data, row)

        return self.data_all
    

    def fill_data_row(self, primary_ids, data, row):
        level_tar = self.headers[0].level if len(self.headers) > 0 else 0
        self.collect(primary_ids, data, row, 0, level_tar)
        return
    

    ###
        # TODO primary_ids stack is not needed, primary_history is enough to perform logic 
        #      (we don't even need to keep a list of primary_history, we could make primary identity as a member of Header)
    ###
    def collect(self, primary_ids, data, row, i, level_tar):

        while i < len(self.headers):

            header = self.headers[i]
            level = header.level

            while len(primary_ids) > 1 and primary_ids[-1].header.level > level:
                primary_ids.pop()
            primary_id = primary_ids[-1]

            if level < level_tar:
                break

            j = None
            value = None

            is_primary_header = header.primary_key == header
            if primary_id.header.level < level and is_primary_header:
                primary_id = PrimaryKey(header)
                primary_ids.append(primary_id)

            
            if header.type == HeaderType.OBJECT:

                value = self.parse_cell_value(row[header.column])
                # if not self.is_empty(value):
                if value is not None:

                    if is_primary_header and primary_id.header.level == header.level and primary_id.identity != value:
                        primary_id.identity = value
                        self.update_primary_history(header, value)
                    
                    if isinstance(data, dict):
                        data[header.name] = value

                    elif isinstance(data, list):
                        item = self.find_primary_item(primary_id, data)
                        if item is None:
                            primary_history_val = self.get_primary_history_value(header)
                            if primary_history_val is not None:
                                temp_primary_id = PrimaryKey(header.primary_key)
                                temp_primary_id.identity= self.get_primary_history_value(header)
                                item = self.find_primary_item(temp_primary_id, data)
                            if item is None:
                                item = {}
                                data.append(item)

                        item[header.name] = value


            elif header.type == HeaderType.DICT:

                subdata = data[header.name] if header.name in data else {}
                value, j = self.collect(primary_ids, subdata, row, i+1, level+1)
                if not self.is_empty(value):

                    if isinstance(data, dict):
                        data[header.name] = value
                    
                    elif isinstance(data, list):
                        item = self.find_primary_item(primary_id, data)
                        if item is None:
                            primary_history_val = self.get_primary_history_value(header)
                            if primary_history_val is not None:
                                temp_primary_id = PrimaryKey(header.primary_key)
                                temp_primary_id.identity = self.get_primary_history_value(header)
                                item = self.find_primary_item(temp_primary_id, data)
                            if item is None:
                                item = {}
                                data.append(item)

                        item[header.name] = value


            else: ## header.type == HeaderType.LIST:


                if isinstance(data, dict):

                    if header.name not in data:
                        data[header.name] = []
                    
                    if len(header.children) > 0:

                        subdata = data[header.name] if header.name in data else {}
                        value, j = self.collect(primary_ids, subdata, row, i+1, level+1)
                        
                    else:
                        value = self.parse_cell_value(row[header.column])
                        if not self.is_empty(value):
                            data[header.name].append(value)
                    


                if isinstance(data, list):

                    item_idx = -1
                    item = self.find_primary_item(primary_id, data, item_idx)

                    ## TODO the logic of making 'item' could  be merged 
                    if len(header.children) > 0:

                        if item is None:
                            primary_history_val = self.get_primary_history_value(header)
                            if primary_history_val is not None:
                                temp_primary_id = PrimaryKey(header.primary_key)
                                temp_primary_id.identity = self.get_primary_history_value(header)
                                item = self.find_primary_item(temp_primary_id, data)

                        if item is None:
                            if len(data) == 0:
                                data.append({})
                            item_idx = 0
                            item = data[-1]

                        allowed_primary_id = primary_id.identity is not None or primary_id.header is None
                        if header.name not in item and primary_id.identity is not allowed_primary_id:
                            item[header.name] = []

                        subdata = item[header.name] if header.name in item else {}
                        value, j = self.collect(primary_ids, subdata, row, i+1, level+1)

                        if self.is_empty(item):
                            data.remove(item)
                            
                    else:

                        value = self.parse_cell_value(row[header.column])
                        if not self.is_empty(value):

                            if item is None:
                                primary_history_val = self.get_primary_history_value(header)
                                if primary_history_val is not None:
                                    temp_primary_id = PrimaryKey(header.primary_key)
                                    temp_primary_id.identity = self.get_primary_history_value(header)
                                    item = self.find_primary_item(temp_primary_id, data)

                            if item is None:
                                if len(data) == 0:
                                    data.append({})
                                item_idx = 0
                                item = data[-1]
                            
                            if header.name not in item and primary_id.identity is not allowed_primary_id:
                                item[header.name] = []
                                
                            item[header.name].append(value)
                    #~
            
            i = i + 1 if j is None else j

        return data, i