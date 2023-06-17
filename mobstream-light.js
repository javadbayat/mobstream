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
}