from enum import Enum
import re


PrefixList = '+'
PrefixGroup = "#"

class HeaderType(Enum):
    ROOT = "root"
    OBJECT = "object"
    DICT = "dict"
    LIST = "list"

class Header:

    def __repr__(self):
        return f"Header(name={self.name}, level={self.level}, type={self.type}, column={self.column})"

    def __init__(self, name, level, header_type, column=None):
        self.name = name
        self.level = level
        self.type = header_type
        self.parent = None
        self.children = []
        self.column = column
        self.primary_key = None
        
    def update_parent(self, parent):
        self.parent = parent
        if parent:
            parent.children.append(self)
            parent.update_type()

    def update_type(self):
        if self.type != HeaderType.LIST and len(self.children) > 0:
            self.type = HeaderType.DICT

    def set_primary_key(self, primary_key):
        self.primary_key = primary_key

    def get_header_path(self):
        path = []
        current = self
        while current and current.type != HeaderType.ROOT:
            path.insert(0, current)
            current = current.parent
        return path
    
    def get_root(self):
        current = self
        while current.parent is not None:
            current = current.parent
        return current

    @classmethod
    def create_root(cls):
        return cls(name="root", level=-1, header_type=HeaderType.ROOT)



def parse_headers(df):

    headers = []
    header_root = Header.create_root()
    parent_stack = [header_root]
    last_level = -1
    primary_key_stack = []
    
    for col_idx, header_str in enumerate(df.columns):

        leading_symbols = ''.join(c for c in header_str if c in (PrefixList, PrefixGroup))
        symbol_len = len(leading_symbols)
        level = symbol_len - 1 if symbol_len > 0 else 0
        
        header_name = header_str.lstrip(PrefixList + PrefixGroup).strip()
        header_name = re.sub(r'\.\d+$', '', header_name)
        
        bIsList = symbol_len > 0 and leading_symbols[0] == PrefixList
        header_type = HeaderType.LIST if bIsList else HeaderType.OBJECT
        
        while parent_stack and parent_stack[-1].level >= level:
            parent_stack.pop()
        parent = parent_stack[-1]
        
        header = Header(name=header_name, level=level, header_type=header_type, column=col_idx)
        header.update_parent(parent)
        headers.append(header)
        parent_stack.append(header)
        
        if last_level < level:
            last_level = level
            primary_key = header if header_type == HeaderType.OBJECT else None
            if primary_key:
                primary_key_stack.append(primary_key)
        else:
            while last_level > level and len(primary_key_stack) > 1:
                last_level -= 1
                primary_key_stack.pop()
        
        header.set_primary_key(primary_key_stack[-1])
    
    return headers, header_root


class PrimaryKey:
    
    def __repr__(self):
        return f"PrimaryKey(header={self.header}, identity={self.identity})"
    
    def __init__(self, header, identity = None):
        self.header = header
        self.identity = identity