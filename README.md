# WordMathCopy

通过直接复制word文档中的MathType公式并黏贴到web中的小工具

内部逻辑为
- 直接从word文档中复制包含数学公式的内容
- 截取系统剪贴板的rtf格式数据进行数学公式解析
- 将数学公式解析成图片文本及方便web传输存储的latex格式公式
- 替换系统剪贴板的rtf格式为html内容并直接在web中使用

使用插件
- [nrtftree](https://github.com/sgolivernet/nrtftree) rtf解析器
- [MTSDKDN](http://www.dessci.com/en/reference/sdk/) MathType SDK
