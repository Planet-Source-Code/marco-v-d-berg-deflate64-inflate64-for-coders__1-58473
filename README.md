<div align="center">

## Deflate64/Inflate64 for coders


</div>

### Description

Deflate and inflate modules.

I saw the article from Jim Reforma about how a deflate algorithm should work and noticed that people asked for working Deflate code. well here it is. This code can be used to create filetypes wich used Deflate as compression algo (ZIP,GZIP,CAB, etc.etc). The only thing you have to do is create the headerdata for these formats.(don't get me wrong but this can be a lot of work(bin there, Done that)).

Don't expect to much of the code in terms of speed since it's all coded in VB, it is pretty fast but not by a long shot as fast as winzip.

If you plan to code a ZIP compliant compressor, then take a look at my Archive Explorer And you will have ALMOST everything you need to know.
 
### More Info
 
call Deflate(ByteArray() As Byte, CompLevel As Integer, Optional Deflate64 As Boolean = False)

ByteArray() = the file to be compressed

Complevel = compression level 0-9

Deflate64(true) = compress in deflate64 format, otherwise in normal deflate format

Inflate(ByteArray() As Byte, Optional UncompressedSize As Long = 1000, Optional ZIP64 As Boolean = False) As Long

ByteArray() = the file to be decompressed

UncompressedSize = if known the decompressed filesize

ZIP64(True) = Decompress in deflate64 format, otherwise in normal deflate format.

The data wich returns in ByteArray() is the compressed or decompressed data.


<span>             |<span>
---                |---
**Submitted On**   |2004-02-09 15:32:08
**By**             |[Marco v/d Berg](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marco-v-d-berg.md)
**Level**          |Intermediate
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Deflate64\_1843611242005\.zip](https://github.com/Planet-Source-Code/marco-v-d-berg-deflate64-inflate64-for-coders__1-58473/archive/master.zip)








