Attribute VB_Name = "sqlite3win32helper"
Option Explicit

' Primary Result Codes
Public Const SQLITE_OK As Long = 0
Public Const SQLITE_ERROR As Long = 1
Public Const SQLITE_INTERNAL As Long = 2
Public Const SQLITE_PERM As Long = 3
Public Const SQLITE_ABORT As Long = 4
Public Const SQLITE_BUSY As Long = 5
Public Const SQLITE_LOCKED As Long = 6
Public Const SQLITE_NOMEM As Long = 7
Public Const SQLITE_READONLY As Long = 8
Public Const SQLITE_INTERRUPT As Long = 9
Public Const SQLITE_IOERR As Long = 10
Public Const SQLITE_CORRUPT As Long = 11
Public Const SQLITE_NOTFOUND As Long = 12
Public Const SQLITE_FULL As Long = 13
Public Const SQLITE_CANTOPEN As Long = 14
Public Const SQLITE_PROTOCOL As Long = 15
Public Const SQLITE_EMPTY As Long = 16
Public Const SQLITE_SCHEMA As Long = 17
Public Const SQLITE_TOOBIG As Long = 18
Public Const SQLITE_CONSTRAINT As Long = 19
Public Const SQLITE_MISMATCH As Long = 20
Public Const SQLITE_MISUSE As Long = 21
Public Const SQLITE_NOLFS As Long = 22
Public Const SQLITE_AUTH As Long = 23
Public Const SQLITE_FORMAT As Long = 24
Public Const SQLITE_RANGE As Long = 25
Public Const SQLITE_NOTADB As Long = 26
Public Const SQLITE_NOTICE As Long = 27
Public Const SQLITE_WARNING As Long = 28
Public Const SQLITE_ROW As Long = 100
Public Const SQLITE_DONE As Long = 101

' Extended Result Codes
Public Const SQLITE_ERROR_MISSING_COLLSEQ As Long = &H101
Public Const SQLITE_ERROR_RETRY As Long = &H201
Public Const SQLITE_ERROR_SNAPSHOT As Long = &H301
Public Const SQLITE_IOERR_READ As Long = &H10A
Public Const SQLITE_IOERR_SHORT_READ As Long = &H20A
Public Const SQLITE_IOERR_WRITE As Long = &H30A
Public Const SQLITE_IOERR_FSYNC As Long = &H40A
Public Const SQLITE_IOERR_DIR_FSYNC As Long = &H50A
Public Const SQLITE_IOERR_TRUNCATE As Long = &H60A
Public Const SQLITE_IOERR_FSTAT As Long = &H70A
Public Const SQLITE_IOERR_UNLOCK As Long = &H80A
Public Const SQLITE_IOERR_RDLOCK As Long = &H90A
Public Const SQLITE_IOERR_DELETE As Long = &HA0A
Public Const SQLITE_IOERR_BLOCKED As Long = &HB0A
Public Const SQLITE_IOERR_NOMEM As Long = &HC0A
Public Const SQLITE_IOERR_ACCESS As Long = &HD0A
Public Const SQLITE_IOERR_CHECKRESERVEDLOCK As Long = &HE0A
Public Const SQLITE_IOERR_LOCK As Long = &HF0A
Public Const SQLITE_IOERR_CLOSE As Long = &H100A
Public Const SQLITE_IOERR_DIR_CLOSE As Long = &H110A
Public Const SQLITE_IOERR_SHMOPEN As Long = &H120A
Public Const SQLITE_IOERR_SHMSIZE As Long = &H130A
Public Const SQLITE_IOERR_SHMLOCK As Long = &H140A
Public Const SQLITE_IOERR_SHMMAP As Long = &H150A
Public Const SQLITE_IOERR_SEEK As Long = &H160A
Public Const SQLITE_IOERR_DELETE_NOENT As Long = &H170A
Public Const SQLITE_IOERR_MMAP As Long = &H180A
Public Const SQLITE_IOERR_GETTEMPPATH As Long = &H190A
Public Const SQLITE_IOERR_CONVPATH As Long = &H1A0A
Public Const SQLITE_IOERR_VNODE As Long = &H1B0A
Public Const SQLITE_IOERR_AUTH As Long = &H1C0A
Public Const SQLITE_IOERR_BEGIN_ATOMIC As Long = &H1D0A
Public Const SQLITE_IOERR_COMMIT_ATOMIC As Long = &H1E0A
Public Const SQLITE_IOERR_ROLLBACK_ATOMIC As Long = &H1F0A
Public Const SQLITE_LOCKED_SHAREDCACHE As Long = &H106
Public Const SQLITE_LOCKED_VTAB As Long = &H206
Public Const SQLITE_BUSY_RECOVERY As Long = &H105
Public Const SQLITE_BUSY_SNAPSHOT As Long = &H205
Public Const SQLITE_CANTOPEN_NOTEMPDIR As Long = &H10E
Public Const SQLITE_CANTOPEN_ISDIR As Long = &H20E
Public Const SQLITE_CANTOPEN_FULLPATH As Long = &H30E
Public Const SQLITE_CANTOPEN_CONVPATH As Long = &H40E
Public Const SQLITE_CANTOPEN_DIRTYWAL As Long = &H50E
Public Const SQLITE_CANTOPEN_SYMLINK As Long = &H60E
Public Const SQLITE_CORRUPT_VTAB As Long = &H10B
Public Const SQLITE_CORRUPT_SEQUENCE As Long = &H20B
Public Const SQLITE_READONLY_RECOVERY As Long = &H108
Public Const SQLITE_READONLY_CANTLOCK As Long = &H208
Public Const SQLITE_READONLY_ROLLBACK As Long = &H308
Public Const SQLITE_READONLY_DBMOVED As Long = &H408
Public Const SQLITE_READONLY_CANTINIT As Long = &H508
Public Const SQLITE_READONLY_DIRECTORY As Long = &H608
Public Const SQLITE_ABORT_ROLLBACK As Long = &H204
Public Const SQLITE_CONSTRAINT_CHECK As Long = &H113
Public Const SQLITE_CONSTRAINT_COMMITHOOK As Long = &H213
Public Const SQLITE_CONSTRAINT_FOREIGNKEY As Long = &H313
Public Const SQLITE_CONSTRAINT_FUNCTION As Long = &H413
Public Const SQLITE_CONSTRAINT_NOTNULL As Long = &H513
Public Const SQLITE_CONSTRAINT_PRIMARYKEY As Long = &H613
Public Const SQLITE_CONSTRAINT_TRIGGER As Long = &H713
Public Const SQLITE_CONSTRAINT_UNIQUE As Long = &H813
Public Const SQLITE_CONSTRAINT_VTAB As Long = &H913
Public Const SQLITE_CONSTRAINT_ROWID As Long = &HA13
Public Const SQLITE_CONSTRAINT_PINNED As Long = &HB13
Public Const SQLITE_NOTICE_RECOVER_WAL As Long = &H11B
Public Const SQLITE_NOTICE_RECOVER_ROLLBACK As Long = &H21B
Public Const SQLITE_WARNING_AUTOINDEX As Long = &H11C
Public Const SQLITE_AUTH_USER As Long = &H117
Public Const SQLITE_OK_LOAD_PERMANENTLY As Long = &H100
Public Const SQLITE_OK_SYMLINK As Long = &H200

' File Open Operation Flags
Public Const SQLITE_OPEN_READONLY As Long = &H1                     ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_READWRITE As Long = &H2                    ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_CREATE As Long = &H4                       ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_DELETEONCLOSE As Long = &H8                ' VFS only
Public Const SQLITE_OPEN_EXCLUSIVE As Long = &H10                   ' VFS only
Public Const SQLITE_OPEN_AUTOPROXY As Long = &H20                   ' VFS only
Public Const SQLITE_OPEN_URI As Long = &H40                         ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_MEMORY As Long = &H80                      ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_MAIN_DB As Long = &H100                    ' VFS only
Public Const SQLITE_OPEN_TEMP_DB As Long = &H200                    ' VFS only
Public Const SQLITE_OPEN_TRANSIENT_DB As Long = &H400               ' VFS only
Public Const SQLITE_OPEN_MAIN_JOURNAL As Long = &H800               ' VFS only
Public Const SQLITE_OPEN_TEMP_JOURNAL As Long = &H1000              ' VFS only
Public Const SQLITE_OPEN_SUBJOURNAL As Long = &H2000                ' VFS only
Public Const SQLITE_OPEN_MASTER_JOURNAL As Long = &H4000            ' VFS only
Public Const SQLITE_OPEN_NOMUTEX As Long = &H8000                   ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_FULLMUTEX As Long = &H10000                ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_SHAREDCACHE As Long = &H20000              ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_PRIVATECACHE As Long = &H40000             ' OK for sqlite3_open_v2()
Public Const SQLITE_OPEN_WAL As Long = &H80000                      ' VFS only
Public Const SQLITE_OPEN_NOFOLLOW As Long = &H1000000               ' OK for sqlite3_open_v2()

' Data Types
Public Const SQLITE_INTEGER As Long = 1
Public Const SQLITE_FLOAT As Long = 2
Public Const SQLITE_TEXT As Long = 3
Public Const SQLITE_BLOB As Long = 4
Public Const SQLITE_NULL As Long = 5

' Special Destructor Behavior Constants
Public Const SQLITE_STATIC As Long = &H0
Public Const SQLITE_TRANSIENT As Long = &HFFFFFFFF

' Text Encodings
Public Const SQLITE_UTF8 As Long = 1
Public Const SQLITE_UTF16LE As Long = 2
Public Const SQLITE_UTF16BE As Long = 3
Public Const SQLITE_UTF16 As Long = 4
Public Const SQLITE_ANY As Long = 5                                 ' Deprecated
Public Const SQLITE_UTF16_ALIGNED As Long = 8

' Function Flags
Public Const SQLITE_DETERMINISTIC As Long = &H800
Public Const SQLITE_DIRECTONLY As Long = &H80000
Public Const SQLITE_SUBTYPE As Long = &H100000
Public Const SQLITE_INNOCUOUS As Long = &H200000

' Device Characteristics
Public Const SQLITE_IOCAP_ATOMIC As Long = &H1
Public Const SQLITE_IOCAP_ATOMIC512 As Long = &H2
Public Const SQLITE_IOCAP_ATOMIC1K As Long = &H4
Public Const SQLITE_IOCAP_ATOMIC2K As Long = &H8
Public Const SQLITE_IOCAP_ATOMIC4K As Long = &H10
Public Const SQLITE_IOCAP_ATOMIC8K As Long = &H20
Public Const SQLITE_IOCAP_ATOMIC16K As Long = &H40
Public Const SQLITE_IOCAP_ATOMIC32K As Long = &H80
Public Const SQLITE_IOCAP_ATOMIC64K As Long = &H100
Public Const SQLITE_IOCAP_SAFE_APPEND As Long = &H200
Public Const SQLITE_IOCAP_SEQUENTIAL As Long = &H400
Public Const SQLITE_IOCAP_UNDELETABLE_WHEN_OPEN As Long = &H800
Public Const SQLITE_IOCAP_POWERSAFE_OVERWRITE As Long = &H1000
Public Const SQLITE_IOCAP_IMMUTABLE As Long = &H2000
Public Const SQLITE_IOCAP_BATCH_ATOMIC As Long = &H4000

' File Locking Levels
Public Const SQLITE_LOCK_NONE As Long = 0
Public Const SQLITE_LOCK_SHARED As Long = 1
Public Const SQLITE_LOCK_RESERVED As Long = 2
Public Const SQLITE_LOCK_PENDING As Long = 3
Public Const SQLITE_LOCK_EXCLUSIVE As Long = 4

' Synchronization Type Flags
Public Const SQLITE_SYNC_NORMAL As Long = &H2
Public Const SQLITE_SYNC_FULL As Long = &H3
Public Const SQLITE_SYNC_DATAONLY As Long = &H10

' Standard File Control Operator Codes
Public Const SQLITE_FCNTL_LOCKSTATE As Long = 1
Public Const SQLITE_FCNTL_GET_LOCKPROXYFILE As Long = 2
Public Const SQLITE_FCNTL_SET_LOCKPROXYFILE As Long = 3
Public Const SQLITE_FCNTL_LAST_ERRNO As Long = 4
Public Const SQLITE_FCNTL_SIZE_HINT As Long = 5
Public Const SQLITE_FCNTL_CHUNK_SIZE As Long = 6
Public Const SQLITE_FCNTL_FILE_POINTER As Long = 7
Public Const SQLITE_FCNTL_SYNC_OMITTED As Long = 8
Public Const SQLITE_FCNTL_WIN32_AV_RETRY As Long = 9
Public Const SQLITE_FCNTL_PERSIST_WAL As Long = 10
Public Const SQLITE_FCNTL_OVERWRITE As Long = 11
Public Const SQLITE_FCNTL_VFSNAME As Long = 12
Public Const SQLITE_FCNTL_POWERSAFE_OVERWRITE As Long = 13
Public Const SQLITE_FCNTL_PRAGMA As Long = 14
Public Const SQLITE_FCNTL_BUSYHANDLER As Long = 15
Public Const SQLITE_FCNTL_TEMPFILENAME As Long = 16
Public Const SQLITE_FCNTL_MMAP_SIZE As Long = 18
Public Const SQLITE_FCNTL_TRACE As Long = 19
Public Const SQLITE_FCNTL_HAS_MOVED As Long = 20
Public Const SQLITE_FCNTL_SYNC As Long = 21
Public Const SQLITE_FCNTL_COMMIT_PHASETWO As Long = 22
Public Const SQLITE_FCNTL_WIN32_SET_HANDLE As Long = 23
Public Const SQLITE_FCNTL_WAL_BLOCK As Long = 24
Public Const SQLITE_FCNTL_ZIPVFS As Long = 25
Public Const SQLITE_FCNTL_RBU As Long = 26
Public Const SQLITE_FCNTL_VFS_POINTER As Long = 27
Public Const SQLITE_FCNTL_JOURNAL_POINTER As Long = 28
Public Const SQLITE_FCNTL_WIN32_GET_HANDLE As Long = 29
Public Const SQLITE_FCNTL_PDB As Long = 30
Public Const SQLITE_FCNTL_BEGIN_ATOMIC_WRITE As Long = 31
Public Const SQLITE_FCNTL_COMMIT_ATOMIC_WRITE As Long = 32
Public Const SQLITE_FCNTL_ROLLBACK_ATOMIC_WRITE As Long = 33
Public Const SQLITE_FCNTL_LOCK_TIMEOUT As Long = 34
Public Const SQLITE_FCNTL_DATA_VERSION As Long = 35
Public Const SQLITE_FCNTL_SIZE_LIMIT As Long = 36
Public Const SQLITE_FCNTL_CKPT_DONE As Long = 37
Public Const SQLITE_FCNTL_RESERVE_BYTES As Long = 38
Public Const SQLITE_FCNTL_CKPT_START As Long = 39

' xAccess VFS Method Flags
Public Const SQLITE_ACCESS_EXISTS As Long = 0
Public Const SQLITE_ACCESS_READWRITE As Long = 1
Public Const SQLITE_ACCESS_READ As Long = 2

' xShmLock VFS Method Flags
Public Const SQLITE_SHM_UNLOCK As Long = 1
Public Const SQLITE_SHM_LOCK As Long = 2
Public Const SQLITE_SHM_SHARED As Long = 4
Public Const SQLITE_SHM_EXCLUSIVE As Long = 8
Public Const SQLITE_SHM_NLOCK As Long = 8

' Configuration Options
Public Const SQLITE_CONFIG_SINGLETHREAD As Long = 1
Public Const SQLITE_CONFIG_MULTITHREAD As Long = 2
Public Const SQLITE_CONFIG_SERIALIZED As Long = 3
Public Const SQLITE_CONFIG_MALLOC As Long = 4
Public Const SQLITE_CONFIG_GETMALLOC As Long = 5
Public Const SQLITE_CONFIG_SCRATCH As Long = 6
Public Const SQLITE_CONFIG_PAGECACHE As Long = 7
Public Const SQLITE_CONFIG_HEAP As Long = 8
Public Const SQLITE_CONFIG_MEMSTATUS As Long = 9
Public Const SQLITE_CONFIG_MUTEX As Long = 10
Public Const SQLITE_CONFIG_GETMUTEX As Long = 11
Public Const SQLITE_CONFIG_LOOKASIDE As Long = 13
Public Const SQLITE_CONFIG_PCACHE As Long = 14
Public Const SQLITE_CONFIG_GETPCACHE As Long = 15
Public Const SQLITE_CONFIG_LOG As Long = 16
Public Const SQLITE_CONFIG_URI As Long = 17
Public Const SQLITE_CONFIG_PCACHE2 As Long = 18
Public Const SQLITE_CONFIG_GETPCACHE2 As Long = 19
Public Const SQLITE_CONFIG_COVERING_INDEX_SCAN As Long = 20
Public Const SQLITE_CONFIG_SQLLOG As Long = 21
Public Const SQLITE_CONFIG_MMAP_SIZE As Long = 22
Public Const SQLITE_CONFIG_WIN32_HEAPSIZE As Long = 23
Public Const SQLITE_CONFIG_PCACHE_HDRSZ As Long = 24
Public Const SQLITE_CONFIG_PMASZ As Long = 25
Public Const SQLITE_CONFIG_STMTJRNL_SPILL As Long = 26
Public Const SQLITE_CONFIG_SMALL_MALLOC As Long = 27
Public Const SQLITE_CONFIG_SORTERREF_SIZE As Long = 28
Public Const SQLITE_CONFIG_MEMDB_MAXSIZE As Long = 29

' Database Connection Configuration Options
Public Const SQLITE_DBCONFIG_MAINDBNAME As Long = 1000
Public Const SQLITE_DBCONFIG_LOOKASIDE As Long = 1001
Public Const SQLITE_DBCONFIG_ENABLE_FKEY As Long = 1002
Public Const SQLITE_DBCONFIG_ENABLE_TRIGGER As Long = 1003
Public Const SQLITE_DBCONFIG_ENABLE_FTS3_TOKENIZER As Long = 1004
Public Const SQLITE_DBCONFIG_ENABLE_LOAD_EXTENSION As Long = 1005
Public Const SQLITE_DBCONFIG_NO_CKPT_ON_CLOSE As Long = 1006
Public Const SQLITE_DBCONFIG_ENABLE_QPSG As Long = 1007
Public Const SQLITE_DBCONFIG_TRIGGER_EQP As Long = 1008
Public Const SQLITE_DBCONFIG_RESET_DATABASE As Long = 1009
Public Const SQLITE_DBCONFIG_DEFENSIVE As Long = 1010
Public Const SQLITE_DBCONFIG_WRITABLE_SCHEMA As Long = 1011
Public Const SQLITE_DBCONFIG_LEGACY_ALTER_TABLE As Long = 1012
Public Const SQLITE_DBCONFIG_DQS_DML As Long = 1013
Public Const SQLITE_DBCONFIG_DQS_DDL As Long = 1014
Public Const SQLITE_DBCONFIG_ENABLE_VIEW As Long = 1015
Public Const SQLITE_DBCONFIG_LEGACY_FILE_FORMAT As Long = 1016
Public Const SQLITE_DBCONFIG_TRUSTED_SCHEMA As Long = 1017

' Authorizer Return Codes
Public Const SQLITE_DENY As Long = 1
Public Const SQLITE_IGNORE As Long = 2

' Authorizer Action Codes
Public Const SQLITE_CREATE_INDEX As Long = 1
Public Const SQLITE_CREATE_TABLE As Long = 2
Public Const SQLITE_CREATE_TEMP_INDEX As Long = 3
Public Const SQLITE_CREATE_TEMP_TABLE As Long = 4
Public Const SQLITE_CREATE_TEMP_TRIGGER As Long = 5
Public Const SQLITE_CREATE_TEMP_VIEW As Long = 6
Public Const SQLITE_CREATE_TRIGGER As Long = 7
Public Const SQLITE_CREATE_VIEW As Long = 8
Public Const SQLITE_DELETE As Long = 9
Public Const SQLITE_DROP_INDEX As Long = 10
Public Const SQLITE_DROP_TABLE As Long = 11
Public Const SQLITE_DROP_TEMP_INDEX As Long = 12
Public Const SQLITE_DROP_TEMP_TABLE As Long = 13
Public Const SQLITE_DROP_TEMP_TRIGGER As Long = 14
Public Const SQLITE_DROP_TEMP_VIEW As Long = 15
Public Const SQLITE_DROP_TRIGGER As Long = 16
Public Const SQLITE_DROP_VIEW As Long = 17
Public Const SQLITE_INSERT As Long = 18
Public Const SQLITE_PRAGMA As Long = 19
Public Const SQLITE_READ As Long = 20
Public Const SQLITE_SELECT As Long = 21
Public Const SQLITE_TRANSACTION As Long = 22
Public Const SQLITE_UPDATE As Long = 23
Public Const SQLITE_ATTACH As Long = 24
Public Const SQLITE_DETACH As Long = 25
Public Const SQLITE_ALTER_TABLE As Long = 26
Public Const SQLITE_REINDEX As Long = 27
Public Const SQLITE_ANALYZE As Long = 28
Public Const SQLITE_CREATE_VTABLE As Long = 29
Public Const SQLITE_DROP_VTABLE As Long = 30
Public Const SQLITE_FUNCTION As Long = 31
Public Const SQLITE_SAVEPOINT As Long = 32
Public Const SQLITE_COPY As Long = 0                                ' Deprecated
Public Const SQLITE_RECURSIVE As Long = 33

' SQL Trace Event Codes
Public Const SQLITE_TRACE_STMT As Long = &H1
Public Const SQLITE_TRACE_PROFILE As Long = &H2
Public Const SQLITE_TRACE_ROW As Long = &H4
Public Const SQLITE_TRACE_CLOSE As Long = &H8

' Run-Time Limit Categories
Public Const SQLITE_LIMIT_LENGTH As Long = 0
Public Const SQLITE_LIMIT_SQL_LENGTH As Long = 1
Public Const SQLITE_LIMIT_COLUMN As Long = 2
Public Const SQLITE_LIMIT_EXPR_DEPTH As Long = 3
Public Const SQLITE_LIMIT_COMPOUND_SELECT As Long = 4
Public Const SQLITE_LIMIT_VDBE_OP As Long = 5
Public Const SQLITE_LIMIT_FUNCTION_ARG As Long = 6
Public Const SQLITE_LIMIT_ATTACHED As Long = 7
Public Const SQLITE_LIMIT_LIKE_PATTERN_LENGTH As Long = 8
Public Const SQLITE_LIMIT_VARIABLE_NUMBER As Long = 9
Public Const SQLITE_LIMIT_TRIGGER_DEPTH As Long = 10
Public Const SQLITE_LIMIT_WORKER_THREADS As Long = 11

' Prepare Flags
Public Const SQLITE_PREPARE_PERSISTENT As Long = &H1
Public Const SQLITE_PREPARE_NORMALIZE As Long = &H2
Public Const SQLITE_PREPARE_NO_VTAB As Long = &H4

' Transaction States
Public Const SQLITE_TXN_NONE As Long = 0
Public Const SQLITE_TXN_READ As Long = 1
Public Const SQLITE_TXN_WRITE As Long = 2

' Virtual Table Scan Flags
Public Const SQLITE_INDEX_SCAN_UNIQUE As Long = 1

' Virtual Table Constraint Operator Codes
Public Const SQLITE_INDEX_CONSTRAINT_EQ As Long = 2
Public Const SQLITE_INDEX_CONSTRAINT_GT As Long = 4
Public Const SQLITE_INDEX_CONSTRAINT_LE As Long = 8
Public Const SQLITE_INDEX_CONSTRAINT_LT As Long = 16
Public Const SQLITE_INDEX_CONSTRAINT_GE As Long = 32
Public Const SQLITE_INDEX_CONSTRAINT_MATCH As Long = 64
Public Const SQLITE_INDEX_CONSTRAINT_LIKE As Long = 65
Public Const SQLITE_INDEX_CONSTRAINT_GLOB As Long = 66
Public Const SQLITE_INDEX_CONSTRAINT_REGEXP As Long = 67
Public Const SQLITE_INDEX_CONSTRAINT_NE As Long = 68
Public Const SQLITE_INDEX_CONSTRAINT_ISNOT As Long = 69
Public Const SQLITE_INDEX_CONSTRAINT_ISNOTNULL As Long = 70
Public Const SQLITE_INDEX_CONSTRAINT_ISNULL As Long = 71
Public Const SQLITE_INDEX_CONSTRAINT_IS As Long = 72
Public Const SQLITE_INDEX_CONSTRAINT_FUNCTION As Long = 150

' Mutex Types
Public Const SQLITE_MUTEX_FAST As Long = 0
Public Const SQLITE_MUTEX_RECURSIVE As Long = 1
Public Const SQLITE_MUTEX_STATIC_MASTER As Long = 2
Public Const SQLITE_MUTEX_STATIC_MEM As Long = 3
Public Const SQLITE_MUTEX_STATIC_MEM2 As Long = 4
Public Const SQLITE_MUTEX_STATIC_OPEN As Long = 4
Public Const SQLITE_MUTEX_STATIC_PRNG As Long = 5
Public Const SQLITE_MUTEX_STATIC_LRU As Long = 6
Public Const SQLITE_MUTEX_STATIC_LRU2 As Long = 7
Public Const SQLITE_MUTEX_STATIC_PMEM As Long = 7
Public Const SQLITE_MUTEX_STATIC_APP1 As Long = 8
Public Const SQLITE_MUTEX_STATIC_APP2 As Long = 9
Public Const SQLITE_MUTEX_STATIC_APP3 As Long = 10
Public Const SQLITE_MUTEX_STATIC_VFS1 As Long = 11
Public Const SQLITE_MUTEX_STATIC_VFS2 As Long = 12
Public Const SQLITE_MUTEX_STATIC_VFS3 As Long = 13

' Status Parameters
Public Const SQLITE_STATUS_MEMORY_USED As Long = 0
Public Const SQLITE_STATUS_PAGECACHE_USED As Long = 1
Public Const SQLITE_STATUS_PAGECACHE_OVERFLOW As Long = 2
Public Const SQLITE_STATUS_SCRATCH_USED As Long = 3
Public Const SQLITE_STATUS_SCRATCH_OVERFLOW As Long = 4
Public Const SQLITE_STATUS_MALLOC_SIZE As Long = 5
Public Const SQLITE_STATUS_PARSER_STACK As Long = 6
Public Const SQLITE_STATUS_PAGECACHE_SIZE As Long = 7
Public Const SQLITE_STATUS_SCRATCH_SIZE As Long = 8
Public Const SQLITE_STATUS_MALLOC_COUNT As Long = 9

' Database Connection Status Parameters
Public Const SQLITE_DBSTATUS_LOOKASIDE_USED As Long = 0
Public Const SQLITE_DBSTATUS_CACHE_USED As Long = 1
Public Const SQLITE_DBSTATUS_SCHEMA_USED As Long = 2
Public Const SQLITE_DBSTATUS_STMT_USED As Long = 3
Public Const SQLITE_DBSTATUS_LOOKASIDE_HIT As Long = 4
Public Const SQLITE_DBSTATUS_LOOKASIDE_MISS_SIZE As Long = 5
Public Const SQLITE_DBSTATUS_LOOKASIDE_MISS_FULL As Long = 6
Public Const SQLITE_DBSTATUS_CACHE_HIT As Long = 7
Public Const SQLITE_DBSTATUS_CACHE_MISS As Long = 8
Public Const SQLITE_DBSTATUS_CACHE_WRITE As Long = 9
Public Const SQLITE_DBSTATUS_DEFERRED_FKS As Long = 10
Public Const SQLITE_DBSTATUS_CACHE_USED_SHARED As Long = 11
Public Const SQLITE_DBSTATUS_CACHE_SPILL As Long = 12

' Prepared Statement Status Parameters
Public Const SQLITE_STMTSTATUS_FULLSCAN_STEP As Long = 1
Public Const SQLITE_STMTSTATUS_SORT As Long = 2
Public Const SQLITE_STMTSTATUS_AUTOINDEX As Long = 3
Public Const SQLITE_STMTSTATUS_VM_STEP As Long = 4
Public Const SQLITE_STMTSTATUS_REPREPARE As Long = 5
Public Const SQLITE_STMTSTATUS_RUN As Long = 6
Public Const SQLITE_STMTSTATUS_MEMUSED As Long = 99

' Checkpoint Mode Values
Public Const SQLITE_CHECKPOINT_PASSIVE As Long = 0
Public Const SQLITE_CHECKPOINT_FULL As Long = 1
Public Const SQLITE_CHECKPOINT_RESTART As Long = 2
Public Const SQLITE_CHECKPOINT_TRUNCATE As Long = 3

' Virtual Table Configuration Options
Public Const SQLITE_VTAB_CONSTRAINT_SUPPORT As Long = 1
Public Const SQLITE_VTAB_INNOCUOUS As Long = 2
Public Const SQLITE_VTAB_DIRECTONLY As Long = 3

' Conflict Resolution Modes
Public Const SQLITE_ROLLBACK As Long = 1
'            SQLITE_IGNORE = 2 -> Authorizer Return Codes
Public Const SQLITE_FAIL As Long = 3
'            SQLITE_ABORT = 4 -> Primary Result Codes
Public Const SQLITE_REPLACE As Long = 5

' Prepared Statement Scan Status Operator Codes
Public Const SQLITE_SCANSTAT_NLOOP As Long = 0
Public Const SQLITE_SCANSTAT_NVISIT As Long = 1
Public Const SQLITE_SCANSTAT_EST As Long = 2
Public Const SQLITE_SCANSTAT_NAME As Long = 3
Public Const SQLITE_SCANSTAT_EXPLAIN As Long = 4
Public Const SQLITE_SCANSTAT_SELECTID As Long = 5

' Serialize Flags
Public Const SQLITE_SERIALIZE_NOCOPY As Long = &H1

' Deserialize Flags
Public Const SQLITE_DESERIALIZE_FREEONCLOSE As Long = 1
Public Const SQLITE_DESERIALIZE_RESIZEABLE As Long = 2
Public Const SQLITE_DESERIALIZE_READONLY As Long = 4

' Win32 Directory Types
Public Const SQLITE_WIN32_DATA_DIRECTORY_TYPE As Long = 1
Public Const SQLITE_WIN32_TEMP_DIRECTORY_TYPE As Long = 2

' Limit Constants
Public Const SQLITE_MAX_LENGTH As Long = 1000000000
Public Const SQLITE_MAX_COLUMN As Long = 2000
Public Const SQLITE_MAX_SQL_LENGTH As Long = 1000000000
Public Const SQLITE_MAX_EXPR_DEPTH As Long = 1000
Public Const SQLITE_MAX_COMPOUND_SELECT As Long = 500
Public Const SQLITE_MAX_VDBE_OP As Long = 25000
Public Const SQLITE_MAX_FUNCTION_ARG As Long = 127
Public Const SQLITE_MAX_ATTACHED As Long = 10
Public Const SQLITE_MAX_VARIABLE_NUMBER As Long = 999
Public Const SQLITE_MAX_PAGE_SIZE As Long = 65536
Public Const SQLITE_MAX_DEFAULT_PAGE_SIZE As Long = 8192
Public Const SQLITE_MAX_PAGE_COUNT As Long = 1073741823
Public Const SQLITE_MAX_LIKE_PATTERN_LENGTH As Long = 50000
Public Const SQLITE_MAX_TRIGGER_DEPTH As Long = 1000
