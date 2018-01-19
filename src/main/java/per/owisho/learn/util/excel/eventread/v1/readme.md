# 通过事件方式读取excel文档

目的：为了减少服务器的内存占用     

缺点：仅支持07以后的excel版本即xlsx格式的excel文件读取      

实现原理：基于excel对OPEN XML格式的支持和SAX方式缓冲解析XML文件

