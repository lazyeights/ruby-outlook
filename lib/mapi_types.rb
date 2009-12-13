module MAPI

  module Types

    # Mapi property types, taken from http://msdn2.microsoft.com/en-us/library/bb147591.aspx.
    # The fields are (data type => [mapi name, variant name, description]).
    PROPERTY_TYPES = {
      0x0000 => ['PT_UNSPECIFICIED', '',  '(Reserved for interface use) type doesnt matter to caller' ],
      0x0001 => ['PT_NULL', 'VT_NULL', 'Null (no valid data)'],
      0x0002 => ['PT_SHORT', 'VT_I2', '16-bit integer (signed)'],
      0x0003 => ['PT_LONG', 'VT_I4', '32-bit integer (signed)'],
      0x0004 => ['PT_FLOAT', 'VT_R4', '32-bit real (floating point)'],
      0x0005 => ['PT_DOUBLE', 'VT_R8', '64-bit real (floating point double)'],
      0x0006 => ['PT_CURRENCY', 'VT_CY', '8-byte integer (scaled by 10,000)'],
      0x0006 => ['PT_APPTIME', '', 'Application time'],
      0x000a => ['PT_ERROR', 'VT_ERROR', 'SCODE value; 32-bit unsigned integer'],
      0x000b => ['PT_BOOLEAN', 'VT_BOOL', 'Boolean'],
      0x000d => ['PT_OBJECT', 'VT_UNKNOWN', 'Data object'],
      0x0014 => ['PT_LONGLONG', '', '8-byte signed integer '],
      0x001e => ['PT_STRING8', 'VT_BSTR', 'String'],
      0x001f => ['PT_UNICODE', 'VT_BSTR', 'String'],
      0x0040 => ['PT_SYSTIME', 'VT_DATE', '8-byte real (date in integer, time in fraction)'],
      0x0048 => ['PT_CLSID', '', 'OLE GUID'],
      0x0102 => ['PT_BINARY', 'VT_BLOB', 'Binary (unknown format)'],
      0x1002 => ['PT_MV_SHORT', 'VT_I2', '16-bit integer (signed)'],
      0x1003 => ['PT_MV_LONG', 'VT_I4', '32-bit integer (signed)'],
      0x1004 => ['PT_MV_FLOAT', 'VT_R4', '32-bit real (floating point)'],
      0x1005 => ['PT_MV_DOUBLE', 'VT_R8', '64-bit real (floating point double)'],
      0x1006 => ['PT_MV_CURRENCY', 'VT_CY', '8-byte integer (scaled by 10,000)'],
      0x1006 => ['PT_MV_APPTIME', '', 'Application time'],
      0x1014 => ['PT_MV_LONGLONG', '', '8-byte signed integer '],
      0x101e => ['PT_MV_STRING8', 'VT_BSTR', 'String'],
      0x101f => ['PT_MV_UNICODE', 'VT_BSTR', 'String'],
      0x1040 => ['PT_MV_SYSTIME', 'VT_DATE', '8-byte real (date in integer, time in fraction)'],
      0x1102 => ['PT_MV_BINARY', 'VT_BLOB', 'Binary (unknown format)']
    }
    
    PROPERTY_TYPES.each { |num, (mapi_name, variant_name, desc)| const_set mapi_name, num }

  end

end
