Attribute VB_Name = "modGetCacheEntry"
'FindFirstUrlCacheEntry: Begins the enumeration of the Internet cache.
'Parameters:
'   lpszUrlSearchPattern:
'   [in] Pointer to a string that contains the source name pattern to search for.
'   This can be set to "cookie:" or "visited:" to enumerate the cookies and URL History
'   entries in the cache. If this parameter is NULL, the function uses *.*.
'   lpFirstCacheEntryInfo:
'   [out] Pointer to an INTERNET_CACHE_ENTRY_INFO structure.
'   lpdwFirstCacheEntryInfoBufferSize:
'   [in, out] Pointer to a variable that specifies the size of the lpFirstCacheEntryInfo
'   buffer, in TCHARs.
'   When the function returns, the variable contains the number of TCHARs copied to the
'   buffer, or the required size needed to retrieve the cache entry, in TCHARs.
'Return Values:
'   Returns a handle that the application can use in the FindNextUrlCacheEntry function
'   to retrieve subsequent entries in the cache.
'   If the function fails, the return value is NULL.

Public Declare Function FindFirstUrlCacheEntry Lib "Wininet.dll" _
                                    Alias "FindFirstUrlCacheEntryA" _
                                    (ByVal lpszUrlSearchPattern As String, _
                                    ByRef lpFirstCacheEntryInfo As Any, _
                                    ByRef lpdwFirstCacheEntryInfoBufferSize As Long) _
                                    As Long
'FindNextUrlCacheEntry: Retrieves the next entry in the Internet cache.
'Parameters:
'   hEnumHandle:
'   [in] Handle to the enumeration obtained from a previous call to FindFirstUrlCacheEntry.
'   lpNextCacheEntryInfo:
'   [out] Pointer to an INTERNET_CACHE_ENTRY_INFO structure that receives information
'   about the cache entry.
'   lpdwNextCacheEntryInfoBufferSize:
'   [in, out] Pointer to a variable that specifies the size of the lpNextCacheEntryInfo
'   buffer, in TCHARs. When the function returns, the variable contains the number of
'   TCHARs copied to the buffer, or the size of the buffer required to retrieve the cache
'   entry, in bytes.
'Return Values
'   Returns TRUE if successful, or FALSE otherwise.
'   To get extended error information, call GetLastError.
'   Possible error values include the following.
'   ERROR_INSUFFICIENT_BUFFER:
'   The size of lpNextCacheEntryInfo as specified by lpdwNextCacheEntryInfoBufferSize is
'   not sufficient to contain all the information.
'   The value returned in lpdwNextCacheEntryInfoBufferSize indicates the buffer size
'   necessary to contain all the information.
'   ERROR_NO_MORE_ITEMS:
'   The enumeration completed.
Public Declare Function FindNextUrlCacheEntry Lib "Wininet.dll" _
                                Alias "FindNextUrlCacheEntryA" _
                                (ByVal hEnumHandle As Long, _
                                ByRef lpNextCacheEntryInfo As Any, _
                                ByRef lpdwNextCacheEntryInfoBufferSize As Long) _
                                As Long
'FindCloseUrlCache:Closes the specified cache enumeration handle.
'Parameters:
'   hEnumHandle:
'   [in] Handle returned by a previous call to the FindFirstUrlCacheEntry function.
'Return Values:
'   Returns TRUE if successful, or FALSE otherwise.
Public Declare Function FindCloseUrlCache Lib "Wininet.dll" _
                                (ByVal hEnumHandle As Long) _
                                As Long

Public Declare Sub CopyMemory Lib "kernel32" _
                                Alias "RtlMoveMemory" _
                                (ByRef Destination As Any, _
                                ByRef Source As Any, _
                                ByVal Length As Long)

Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
                                (ByVal Result As String, _
                                ByVal lpPointer As Long) _
                                As Long

Public Declare Function lstrlen Lib "kernel32" _
                                Alias "lstrlenA" _
                                (ByVal lpPointer As Any) _
                                As Long
                                
Public Declare Function LocalAlloc Lib "kernel32" _
                                (ByVal uFlags As Long, _
                                ByVal uBytes As Long) _
                                As Long

Public Declare Function LocalFree Lib "kernel32" _
                                (ByVal hMem As Long) _
                                As Long

Public Type FILETIME
   dwLowDateTime        As Long
   dwHighDateTime       As Long
End Type

'typedef struct _INTERNET_CACHE_ENTRY_INFO
'{DWORD     dwStructSize;
'LPTSTR     lpszSourceUrlName;
'LPTSTR     lpszLocalFileName;
'DWORD      CacheEntryType;
'DWORD      dwUseCount;
'DWORD      dwHitRate;
'DWORD      dwSizeLow;
'DWORD      dwSizeHigh;
'FILETIME   LastModifiedTime;
'FILETIME   ExpireTime;
'FILETIME   LastAccessTime;
'FILETIME   LastSyncTime;
'LPBYTE     lpHeaderInfo;
'DWORD      dwHeaderInfoSize;
'LPTSTR     lpszFileExtension;
'union
'{DWORD     dwReserved;
' DWORD     dwExemptDelta;
'};
'} INTERNET_CACHE_ENTRY_INFO, *LPINTERNET_CACHE_ENTRY_INFO;

'dwStructSize:          'Size of this structure, in bytes. This value can be used to help determine the version of the cache system.
'lpszSourceUrlName:     'Pointer to a null-terminated string that contains the URL name. The string occupies the memory area at the end of this structure.
'lpszLocalFileName:     'Pointer to a null-terminated string that contains the local file name. The string occupies the memory area at the end of this structure.
'CacheEntryType:        'Cache type bitmask. Currently, the cache entry type value of resources from the Internet is equal to zero. For History and Cookie entries, the cache entry type is a combination of two values. One value determines how the cache entry is handled; the second value indicates what is being cached.
'dwUseCount:            'Current user count of the cache entry.
'dwHitRate:             'Number of times the cache entry was retrieved.
'dwSizeLow:             'Low-order portion of the file size, in TCHARs.
'dwSizeHigh:            'High-order portion of the file size, in TCHARs.
'LastModifiedTime:      'FILETIME structure that contains the last modified time of this URL, in Greenwich mean time format.
'ExpireTime:            'FILETIME structure that contains the expiration time of this file, in Greenwich mean time format.
'LastAccessTime:        'FILETIME structure that contains the last accessed time, in Greenwich mean time format.
'LastSyncTime:          'FILETIME structure that contains the last time the cache was synchronized.
'lpHeaderInfo:          'Pointer to a buffer that contains the header information. The buffer occupies the memory at the end of this structure.
'dwHeaderInfoSize:      'Size of the lpHeaderInfo buffer, in TCHARs.
'lpszFileExtension:     'Pointer to a string that contains the file extension used to retrieve the data as a file. The string occupies the memory area at the end of this structure.
'dwReserved:            'Reserved. Must be zero.
'dwExemptDelta:         'Exemption time from the last accessed time, in seconds.

Public Type INTERNET_CACHE_ENTRY_INFO
   dwStructSize             As Long
   lpszSourceUrlName        As Long
   lpszLocalFileName        As Long
   CacheEntryType           As Long
   dwUseCount               As Long
   dwHitRate                As Long
   dwSizeLow                As Long
   dwSizeHigh               As Long
   LastModifiedTime         As FILETIME
   ExpireTime               As FILETIME
   LastAccessTime           As FILETIME
   LastSyncTime             As FILETIME
   lpHeaderInfo             As Long
   dwHeaderInfoSize         As Long
   lpszFileExtension        As Long
   dwReserved               As Long
   dwExemptDelta            As Long
End Type

Public Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const NORMAL_CACHE_ENTRY = &H1
'LMEM_FIXED = Allocates fixed memory. This flag cannot be combined with the LMEM_MOVEABLE
'or LMEM_DISCARDABLE flag. The return value is a pointer to the memory block.
'To access the memory, the calling process simply casts the return value to a pointer.
Public Const LMEM_FIXED = &H0

'DeleteUrlCacheEntry:
'Removes the file associated with the source name from the cache, if the file exists.
'   lpszUrlName:
'   [in] Pointer to a string that contains the name of the source corresponding to the _
'   cache entry.
'Returns TRUE if successful, or FALSE otherwise. To get extended error information,
'call GetLastError.
'Possible error values include:
'   ERROR_ACCESS_DENIED :
'   The file is locked or in use. The entry will be marked and will be deleted when the
'   file is unlocked.
'   ERROR_FILE_NOT_FOUND :
'   The file is not in the cache.
Public Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
                                    Alias "DeleteUrlCacheEntryA" _
                                    (ByVal lpszUrlName As String) _
                                    As Long
                                    
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_FILE_NOT_FOUND = 2&

