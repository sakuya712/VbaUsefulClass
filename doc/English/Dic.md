
Contents
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [Dic Class](#dic-class)
  - [Methods](#methods)
    - [Add](#add)
    - [AddIf](#addif)
    - [Exists](#exists)
    - [Items](#items)
    - [Key](#key)
    - [Remove](#remove)
    - [SetCompareMode](#setcomparemode)
  - [Properties](#properties)
    - [Count](#count)
    - [DuplicateCollection](#duplicatecollection)
    - [Items](#items-1)
    - [Keys](#keys)

<!-- /code_chunk_output -->

# Dic Class

There is a Dictionary object (associative array) in VBA, but it is troublesome that intellisense cannot be used unless it is referenced.  
Also, since Dictionary can judge whether it is duplicated, you can use it to create a list that does not duplicate, a list that puts out duplicates, and so on.   

## Methods

### Add

\<Argument> Key, element  
\<Return value> None  

- Add an element associated with the key just like a regular Dictionary

### AddIf

\<Argument> Key, element, [DuplicateMode enum]  
\<Return value> None  

- Check if the key is duplicated and add it.
- If there are duplicates, the process determined by the DuplicateMode enumeration type is performed.
- If the 3rd argument is omitted, dmNothing will be processed.
- Duplicate list can be obtained from DuplicateCollection property

**DuplicateMode**
|  Element name           |  Description  |
| ------------------| --------|
|  dmNothing          |  Do nothing without adding|
|  dmOverwrite       |  Overwrite the element|
|  dmSetDuplicateCollection       |  Add to duplicate collection without adding|
|  dmSetBothDuplicateCollection          |  Add to duplicate collection without adding. Add the duplicated side that has already been set|

### Exists

\<Argument> Key  
\<Return value> Boolean value 

- Returns True if the key has already been used


### Items

\<Argument> Key  
\<Return value> element

- Returns the element associated with the key
- Default member
- Read-only unlike the original dictionary

### Key
\<Argument> Key to change, new key  
\<Return value> None

- Change new key (element does not change)

### Remove
\<Argument> Key to delete  
\<Return value> None

- Delete the element

### SetCompareMode

\<Argument> VbCompareMethod enum  
\<Return value> None 

- You can set the comparison mode when comparing string keys.
- It is almost the same as Dictionary, but the setting of vbUseCompareOption is disabled.
- At the time of the constructor, it is the setting of vbTextCompare.


## Properties

### Count

\<Argument> None  
\<Return value> Element count 

- Returns the number of elements

### DuplicateCollection

\<Argument> None  
\<Return value> Collection of duplicate elements  

- If AddIf is set to collect duplicates, the collection will be returned.  
- It is convenient to use it to correct only duplicates or to teach users.

### Items

\<Argument> None  
\<Return value> Array of elements

- Returns the entire contents of the element
- The first element you put in is the index number first

### Keys

\<Argument> None  
\<Return value> Array of keys

- Returns the entire key string
- The first element you put in is the index number first