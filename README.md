# 4d-plugin-xlnt
4D implementation of [xlnt](https://github.com/tfussell/xlnt)

## Syntax (draft)

```
obj:=xlnt IMPORT WORKBOOK (path{;password})
```

Parameter|Type|Description
------------|------------|----
path|TEXT|
password|TEXT|
obj|TEXT|``{"id":id, "class":"workbook"}``

```
xlnt EXPORT WORKBOOK (obj;path{;password})
```

Parameter|Type|Description
------------|------------|----
obj|TEXT|``{"id":id, "class":"workbook"}``
path|TEXT|
password|TEXT|

```
result:=xlnt Workbook from blob (data{;password})

```

Parameter|Type|Description
------------|------------|----
data|BLOB|
password|TEXT|
obj|TEXT|``{"id":id, "class":"workbook"}``

```
xlnt WORKBOOK TO BLOB (obj;data{;password})
```

Parameter|Type|Description
------------|------------|----
obj|TEXT|``{"id":id, "class":"workbook"}``
data|BLOB|
password|TEXT|

```
xlnt CLEAR (obj)
```

Parameter|Type|Description
------------|------------|----
obj|TEXT|``{"id":id, "class":"workbook"}``

---

#### Properties of workbook

``base_date`` (``windows_1900`` or ``mac_1904``)  
``title`` (null if !has_title)  
``sheet_count``  
``active_sheet`` (title)  
``sheet_titles[]``  
``core_properties.category``  
``core_properties.content_status``  
``core_properties.created``  
``core_properties.creator``  
``core_properties.description``  
``core_properties.identifier``  
``core_properties.language``  
``core_properties.last_modified_by``  
``core_properties.last_printed``  
``core_properties.modified``  
``core_properties.revision``  
``core_properties.subject``  
``core_properties.title``  
``core_properties.version``  
  
