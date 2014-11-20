[:arrow_backward:](RTD-Examples.md)

# Type mapping

Following table should be used as a guide for:
- data type conversion
- column type definitions (e.g. for `qTable` function)


```
| q type         | qXL string | VBA type  |
|----------------|------------|-----------|
| boolean        | b          | Boolean   |
| byte           | x          | Byte      |
| short          | h          | Integer   |
| int            | i          | Long      |
| long           | j          | LongLong* |
| real           | e          | Single    |
| float (double) | f          | Double    |
| char           | c          | String    |
| symbol         | s          | String    |
| month          | m          | Date      |
| date           | d          | Date      |
| datetime       | z          | Date      |
| minute         | u          | Date      |
| second         | v          | Date      |
| time           | t          | Date      |
| timestamp      | p          | Date      |
| timespan       | n          | Date      |
```

\* Due to the type `LongLong` being exclusive to 64-bit Excel platforms, in `qXL` for 32-bit Excel the long type is returned as `Double`.

Null types are returned as an empty string.
