// The Magical Output Binary Stream (MOBS) Library - Light Version

function MOBStream(seedPath, seedOffset) {
    var m_stream = new ActiveXObject("ADODB.Stream");
    var m_byteValueMap = new ActiveXObject("Scripting.Dictionary");
    var m_currentFilePath = "";
    
    var adTypeBinary = 1;
    
    if (!seedPath)
        seedPath = "mobstream-seed";
    
    var seedStream = new ActiveXObject("ADODB.Stream");
    seedStream.Open();
    seedStream.Type = adTypeBinary;
    seedStream.LoadFromFile(seedPath);
    
    if (seedOffset)
        seedStream.Read(seedOffset);
    
    for (var b = 0; b < 256; b++)
        m_byteValueMap(b) = seedStream.Read(1);
    
    seedStream.Close();
    
    this.Open = function(filePath) {
        if (!filePath)
            throw new Error("MOBStream.Open: Not enough arguments");
        
        if (m_currentFilePath)
            throw new Error("MOBStream.Open: An open file is currently held");
        
        m_stream.Open();
        m_stream.Type = adTypeBinary;
        m_currentFilePath = filePath;
    };
    
    this.Write = function(data) {
        if (!data)
            throw new Error("MOBStream.Write: Not enough arguments");
        
        if (!m_currentFilePath)
            throw new Error("MOBStream.Write: There is currently no open file to write into");
        
        for (var e = new Enumerator(data); !e.atEnd(); e.moveNext())
            m_stream.Write(m_byteValueMap(e.item()));
    };
    
    this.Close = function() {
        if (!m_currentFilePath)
            throw new Error("MOBStream.Close: There is currently no open file to close");
        
        m_stream.SaveToFile(m_currentFilePath);
        m_stream.Close();
        m_currentFilePath = "";
    };
    
    this.GetSize = function() {
        return m_stream.Size;
    };
    
    this.toString = function() {
        return "[object MOBStream]";
    };
}

function Zipper(mobs) {
    if (!(mobs instanceof MOBStream))
        throw new Error("Zipper: The given argument is not a valid MOBStream instance");
    
    var m_mobs = mobs;
    var m_shell = new ActiveXObject("Shell.Application");
    var m_fso = new ActiveXObject("Scripting.FileSystemObject");
    var m_services = GetObject("winmgmts:");
    var m_zipFolder = null, m_zipFile = null;
    
    var m_eocd = [ 0x50, 0x4B, 0x05, 0x06 ];
    for (var i = 0; i < 18; i++)
        m_eocd.push(0);
    
    var ZAF_DELETE_AFTER_ARCHIVING = 1;
    var ForAppending = 8;
    var SLEEPER_WQL = 'SELECT * FROM __InstanceModificationEvent WHERE TargetInstance ISA "Win32_LocalTime"';
    
    this.Open = function(zipFilePath) {
        if (!zipFilePath)
            throw new Error("Zipper.Open: Not enough arguments");
        
        if (m_zipFile)
            throw new Error("Zipper.Open: An open zip file is currently held");
        
        if (!m_fso.FileExists(zipFilePath)) {
            m_mobs.Open(zipFilePath);
            m_mobs.Write(m_eocd);
            m_mobs.Close();
        }
        
        m_zipFolder = m_shell.NameSpace(zipFilePath);
        m_zipFile = m_fso.GetFile(zipFilePath);
    };
    
    this.Add = function(zipItem, shellFlags, flags) {
        if (!zipItem)
            throw new Error("Zipper.Add: Not enough arguments");
        
        if (!shellFlags)
            shellFlags = 0;
        
        if (!flags)
            flags = 0;
        
        if (!m_zipFolder)
            throw new Error("Zipper.Add: There is currently no open zip file to add items to");
        
        if (flags & ZAF_DELETE_AFTER_ARCHIVING)
            m_zipFolder.MoveHere(zipItem, shellFlags);
        else
            m_zipFolder.CopyHere(zipItem, shellFlags);
    };
    
    function Check() {
        if (!m_zipFile)
            throw new Error("Zipper.Poll: There is currently no open zip file to poll");
        
        try {
            m_zipFile.OpenAsTextStream(ForAppending).Close();
            return true;
        }
        catch (err) {
            if (err.number == -2146828218)
                return false;
            
            throw err;
        }
    }
    
    this.Poll = function(timeout) {
        if (!timeout)
            return Check();
        
        if ((typeof timeout != "number") || (timeout < 0))
            throw new Error("Zipper.Poll: Invalid timeout argument");
        
        var sleeper = m_services.ExecNotificationQuery(SLEEPER_WQL);
        var period = 0;
        
        do {
            sleeper.NextEvent();
            
            if (Check())
                return true;
            
            period += 1000;
        } while (period < timeout);
        
        return false;
    };
    
    this.Close = function() {
        if (!m_zipFile)
            throw new Error("Zipper.Close: There is currently no open zip file to close");
        
        m_zipFile = m_zipFolder = null;
    };
    
    this.toString = function() {
        return "[object Zipper]";
    };
}