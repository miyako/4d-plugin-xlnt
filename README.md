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
``title`` (``null`` if !has_title)  
``sheet_count``  
``active_sheet`` (title)  
``sheet_titles[]``  
``core_properties.category``  
``core_properties.content_status``  
``core_properties.created``  
``core_properties.creator``  
``core_properties.description``  
``core_properties.identifier``  
``core_properties.keywords``  
``core_properties.language``  
``core_properties.last_modified_by``  
``core_properties.last_printed``  
``core_properties.modified``  
``core_properties.revision``  
``core_properties.subject``  
``core_properties.title``  
``core_properties.version``  
  
  ---
  
```
xlnt SET VALUES (obj;values)
xlnt GET VALUES (obj;values)
```

Parameter|Type|Description
------------|------------|----
obj|TEXT|``{"id":id, "class":"workbook", target:name}``
values|TEXT|``value`` or ``[value, value,...]``

``target`` should be ``sheet`` or ``range``  
``obj.range`` is ``named_range``  
``obj.sheet`` is ``sheet_by_title``  

if both are omitted, ``active_sheet`` is implied.  

if ``obj.range`` is specified, ``values`` should be a single ``value`` object.

otherwise, it can be an array of ``value`` objects.   

``value`` object should be ``{location:content}``  

``location`` should be ``cell`` or ``range``  

``cell`` can be ``{"column":col, "row":row}``, ``[col, row]``, ``cell_reference``, or ``null`` (default=A1).  
  
``cell_reference`` is a a cell coordinate (e.g. B12, $B$12).  

``content`` can be ``value``, ``formula``, ``comment``, ``date``, ``time`` or ``datetime``.  

``value`` can be ``null``, boolean, string or number.  

``date`` can be number (days_since_base_year), ``{"year":year, "month":month, "day":day}`` or "today".  

``time`` can be	string, number (fraction of a day where ``0.5`` is noon), ``{"hour":hour, "minute":minute, "second":second, "microsecond":microsecond}`` or 	"now"  

``datetime`` can be ISO string, number (integer=date, fractional=time), ``{"year":year, "month":month, "day":day, "hour":hour, "minute":minute, "second":second, "microsecond":microsecond}``, "now" or "today"

**TODO**

``hyperlink`` as ``value``. ``"hyperlink":url`` or ``"hyperlink":{"cell":cell}`` or ``"hyperlink":{"url":url, "display":display}``  

## Examples

### Set cell value

```
$object:=xlnt IMPORT WORKBOOK ($path;$password)

$o:=JSON Parse($object)

OB SET($o;"sheet";"Sheet1")  //target:sheet

ARRAY OBJECT($values;5)

  //cell can be cell_reference
OB SET($values{1};"cell";"A2";"value";"A2")

  //cell can be {"column":col, "row":row}
C_OBJECT($cells)
OB SET($cells;"column";3;"row";3)
OB SET($values{2};"cell";$cells;"value";"3,3")

  //cell can be [col, row]
ARRAY LONGINT($n;2)
$n{1}:=2
$n{2}:=2
OB SET ARRAY($values{3};"cell";$n)
OB SET($values{3};"value";"2,2")

  //cell can be null
OB SET NULL($values{4};"cell")
OB SET($values{4};"value";"null")

OB SET($values{5};"cell";"D4";"value";"D4")

xlnt SET VALUES (JSON Stringify($o);JSON Stringify array($values))
```

<img width="500" alt="2018-03-08 16 45 56" src="https://user-images.githubusercontent.com/1725068/37139163-341e6b68-22f0-11e8-8dad-c8098c09bc25.png">

### Set date value

```
$object:=xlnt IMPORT WORKBOOK ($path;$password)

$o:=JSON Parse($object)

OB SET($o;"sheet";"Sheet1")  //target:sheet

ARRAY OBJECT($values;3)

OB SET($values{1};"cell";"A1";"date";"today")
C_OBJECT($date)
OB SET($date;"year";2112;"month";9;"day";3)
OB SET($values{2};"cell";"A2";"date";$date)
OB SET($values{3};"cell";"A3";"date";44444)

xlnt SET VALUES (JSON Stringify($o);JSON Stringify array($values))
```

### Set time value

```
$object:=xlnt IMPORT WORKBOOK ($path;$password)

$o:=JSON Parse($object)

OB SET($o;"sheet";"Sheet1")  //target:sheet

ARRAY OBJECT($values;4)

OB SET($values{1};"cell";"A1";"time";"now")
OB SET($values{2};"cell";"A2";"time";0.5)  
C_OBJECT($time)
OB SET($time;"hour";12;"minute";34;"second";56;"microsecond";78)
OB SET($values{3};"cell";"A3";"time";$time)
OB SET($values{4};"cell";"A4";"time";"12:34:56")

xlnt SET VALUES (JSON Stringify($o);JSON Stringify array($values))
```

### Set datetime value

```
$object:=xlnt IMPORT WORKBOOK ($path;$password)

$o:=JSON Parse($object)

OB SET($o;"sheet";"Sheet1")  //target:sheet

ARRAY OBJECT($values;5)

OB SET($values{1};"cell";"A1";"datetime";"today")
C_OBJECT($datetime)
OB SET($datetime;"year";2112;"month";9;"day";3;"hour";12;"minute";34;"second";56;"microsecond";78)
OB SET($values{2};"cell";"A2";"datetime";$datetime)
OB SET($values{3};"cell";"A3";"datetime";44444.5)
OB SET($values{4};"cell";"A4";"datetime";"now")  //same as today
OB SET($values{5};"cell";"A5";"datetime";string(Current date;ISO date;current time))  //same as today

xlnt SET VALUES (JSON Stringify($o);JSON Stringify array($values))
```

### Set comment 

```
$object:=xlnt IMPORT WORKBOOK ($path;$password)

$o:=JSON Parse($object)

OB SET($o;"sheet";"Sheet1")  //target:sheet

ARRAY OBJECT($values;2)

OB SET($values{1};"cell";"A1";"comment";"comment")
OB SET($values{2};"cell";"A1";"value";"value")

xlnt SET VALUES (JSON Stringify($o);JSON Stringify array($values))
```
