# 使用 POI处理 Word 文件
POI 版本： 4.1.2

> **注**： 下面这些操作都是基于 ooxml 规范来进行实现的，因此只能作用在 docx 文档上。


# 测试环境依赖
```java
    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-ooxml</artifactId>
      <version>4.1.2</version>
    </dependency>

    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>ooxml-schemas</artifactId>
      <version>1.4</version>
    </dependency>

    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-scratchpad</artifactId>
      <version>4.1.2</version>
    </dependency>

    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-examples</artifactId>
      <version>4.1.2</version>
    </dependency>

    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi</artifactId>
      <version>4.1.2</version>
    </dependency>
```



# 样式属性封装
## 字体样式
```java
/**
 * @author Zhao Wen
 **/


public class FontStyle extends Style{
    private static final long serialVersionUID = -2234269425804504396L;

    // 字体大小
    private int fontSize = FontSizeConst.XIAOSI;

    // 字体样式
    private String fontFamily = FontFamilyConst.SONGTI;

    // 字体是否为粗体
    private Boolean isBold = false;

    // 字体是否为斜体
    private Boolean isItalic = false;

    // 字体颜色 例如："red"
    private String fontColor = ColorConst.BLANCK;

    public static FontBuilder builder(){
        return new FontBuilder();
    }

    public FontStyle() {

    }

    public FontStyle(int fontSize, String fontFamily) {
        this.fontSize = fontSize;
        this.fontFamily = fontFamily;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public Boolean getBold() {
        return isBold;
    }

    public void setBold(Boolean bold) {
        isBold = bold;
    }

    public Boolean getItalic() {
        return isItalic;
    }

    public void setItalic(Boolean italic) {
        isItalic = italic;
    }

    public String getFontColor() {
        return fontColor;
    }

    public void setFontColor(String fontColor) {
        this.fontColor = fontColor;
    }

    // 设定一个Sytle的builder构造器
    public static class FontBuilder{
        private FontStyle style;

        public FontBuilder() {
            this.style = new FontStyle();
        }

        // 设置各种style的构造器
        public FontBuilder fontSize(int size){
            style.setFontSize(size);
            return this;
        }

        public FontBuilder fontFamily(String family){
            style.setFontFamily(family);
            return this;
        }

        public FontBuilder fontColor(String color){
            style.setFontColor(color);
            return this;
        }

        public FontBuilder buildBold(){
            style.setBold(true);
            return this;
        }

        public FontBuilder buildItalic(){
            style.setItalic(true);
            return this;
        }

        public FontStyle build(){
            return this.style;
        }
    }

}
```

几个常量类：
1. 颜色常量，存储16进制颜色值
```java
/**
 * @author Zhao Wen
 **/
public class ColorConst {
    public static final String BLANCK = "000000";
    public static final String RED = "FF0000";
}
```

2. 字体常量，word 中是什么名字就写入什么名字，中文也好英文也好
```java
/**
 * @author Zhao Wen
 **/
public class FontFamilyConst {
    public static final String TIMENEWROMAN = "Times New Roman";
    public static final String LISHU = "隶书";
    public static final String HEITI = "黑体";
    public static final String SONGTI = "宋体";
    public static final String FANGSONG = "仿宋_GB2312";

    /**
     * 辅助yml文件配置查找字体样式
     * @param flag
     * @return
     */
    public String selectFZ(int flag){
        switch (flag){
            case 1:
                return TIMENEWROMAN;
            case 2:
                return LISHU;
            case 3:
                return HEITI;
            case 4:
                return SONGTI;
            case 5:
                return FANGSONG;
            default: return SONGTI;
        }
    }
}
```


3. 字体大小常量，下面这些大小都是经过检验的
```java
/**
 * @author Zhao Wen
 **/
public class FontSizeConst {
    public static final int XIAOCHU = 72;
    public static final int YIHAO = 52;
    public static final int ERHAO = 44;
    public static final int XIAOER = 36; 
    public static final int XIAOSI = 24; 
    public static final int SIHAO = 28; 
    public static final int WUHAO = 21; 
}
```



## 段落样式
```java
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;

/**
 * @author Zhao Wen
 **/

public class ParagraphStyle extends Style {
    /**
     * 对齐方式 对应 align
     */
    private STJc.Enum align = STJc.LEFT;

    /**
     * 缩进-左侧 对应 LeftChars
     */
    private int indentLeftChars = 0;

    /**
     * 缩进-右侧 对应 RightChars
     */
    private int indentRightChars = 0;

    /**
     * 缩进-悬挂 对应 HangingChars
     */
    private int indentHangingChars = 0;

    private int indentHanging = 0;

    /**
     * 缩进-首行 对应 FirstLineChars
     */
    private int indentFirstLineChars = 0;

    private int indentFirstLine = 0;

    /**
     * 间距-段前 对应 BeforeLines
     */
    private int spacingBeforeLines = 0;

    private int spacingBefore = 0;
    /**
     * 间距-段后 对应 AfterLines
     */
    private int spacingAfterLines = 0;

    private int spacingAfter = 0;
    /**
     * 行距 间距 数值 对应 Line
     */
    private int spacing = 0;

    /**
     * 行距间距对齐规则 对应 LineRule
     */
    private STLineSpacingRule.Enum spacingRule = STLineSpacingRule.AUTO;

    /**
     * 增加编号样式
     */
    private FontStyle glyphStyle;
    private long numId = -1;
    private long lvl = -1;

    // 默认字体样式
    private FontStyle defaultTextStyle;

    private ParagraphStyle(ParagraphBuilder builder) {
        this.align = builder.align;
        this.indentLeftChars = builder.indentLeftChars;
        this.indentRightChars = builder.indentRightChars;
        this.indentHangingChars = builder.indentHangingChars;
        this.indentHanging = builder.indentHanging;
        this.indentFirstLineChars = builder.indentFirstLineChars;
        this.indentFirstLine = builder.indentFirstLine;
        this.spacingBeforeLines = builder.spacingBeforeLines;
        this.spacingAfterLines = builder.spacingAfterLines;
        this.spacing = builder.spacing;
        this.spacingRule = builder.spacingRule;
        this.glyphStyle = builder.glyphStyle;
        this.numId = builder.numId;
        this.lvl = builder.lvl;
        this.defaultTextStyle = builder.defaultTextStyle;
        this.spacingBefore = builder.spacingBefore;
        this.spacingAfter = builder.spacingAfter;
    }

    // 清除缩进信息
    public void cleanAllIndent(){
        this.indentLeftChars = 0;
        this.indentRightChars = 0;
        this.indentHangingChars = 0;
        this.indentHanging = 0;
        this.indentFirstLineChars = 0;
        this.indentFirstLine = 0;
    }

    public STJc.Enum getAlign() {
        return align;
    }

    public void setAlign(STJc.Enum align) {
        this.align = align;
    }

    public int getIndentLeftChars() {
        return indentLeftChars;
    }

    public void setIndentLeftChars(int indentLeftChars) {
        this.indentLeftChars = indentLeftChars;
    }

    public int getIndentRightChars() {
        return indentRightChars;
    }

    public void setIndentRightChars(int indentRightChars) {
        this.indentRightChars = indentRightChars;
    }

    public int getIndentHangingChars() {
        return indentHangingChars;
    }

    public void setIndentHangingChars(int indentHangingChars) {
        this.indentHangingChars = indentHangingChars;
    }

    public int getIndentFirstLineChars() {
        return indentFirstLineChars;
    }

    public void setIndentFirstLineChars(int indentFirstLineChars) {
        this.indentFirstLineChars = indentFirstLineChars;
    }

    public int getSpacingBeforeLines() {
        return spacingBeforeLines;
    }

    public void setSpacingBeforeLines(int spacingBeforeLines) {
        this.spacingBeforeLines = spacingBeforeLines;
    }

    public int getSpacingAfterLines() {
        return spacingAfterLines;
    }

    public void setSpacingAfterLines(int spacingAfterLines) {
        this.spacingAfterLines = spacingAfterLines;
    }

    public int getSpacing() {
        return spacing;
    }

    public void setSpacing(int spacing) {
        this.spacing = spacing;
    }

    public static long getSerialVersionUID() {
        return serialVersionUID;
    }

    public STLineSpacingRule.Enum getSpacingRule() {
        return spacingRule;
    }

    public void setSpacingRule(STLineSpacingRule.Enum spacingRule) {
        this.spacingRule = spacingRule;
    }

    public FontStyle getGlyphStyle() {
        return glyphStyle;
    }

    public void setGlyphStyle(FontStyle glyphStyle) {
        this.glyphStyle = glyphStyle;
    }

    public long getNumId() {
        return numId;
    }

    public void setNumId(long numId) {
        this.numId = numId;
    }

    public long getLvl() {
        return lvl;
    }

    public void setLvl(long lvl) {
        this.lvl = lvl;
    }

    public FontStyle getDefaultTextStyle() {
        return defaultTextStyle;
    }

    public void setDefaultTextStyle(FontStyle defaultTextStyle) {
        this.defaultTextStyle = defaultTextStyle;
    }

    public int getSpacingBefore() {
        return spacingBefore;
    }

    public void setSpacingBefore(int spacingBefore) {
        this.spacingBefore = spacingBefore;
    }

    public int getSpacingAfter() {
        return spacingAfter;
    }

    public void setSpacingAfter(int spacingAfter) {
        this.spacingAfter = spacingAfter;
    }

    public int getIndentHanging() {
        return indentHanging;
    }

    public void setIndentHanging(int indentHanging) {
        this.indentHanging = indentHanging;
    }

    public int getIndentFirstLine() {
        return indentFirstLine;
    }

    public void setIndentFirstLine(int indentFirstLine) {
        this.indentFirstLine = indentFirstLine;
    }

    public static ParagraphBuilder builder(){
        return new ParagraphBuilder();
    }

    public static class ParagraphBuilder{
        private STJc.Enum align = STJc.LEFT;
        private int indentLeftChars = 0;
        private int indentRightChars = 0;
        private int indentHangingChars = 0;
        private int indentHanging = 0;
        private int indentFirstLineChars = 0;
        private int indentFirstLine = 0;
        private int spacingBeforeLines = 0;
        private int spacingBefore = 0;
        private int spacingAfterLines = 0;
        private int spacingAfter = 0;
        private int spacing = 0;
        private STLineSpacingRule.Enum spacingRule = STLineSpacingRule.AUTO;
        private FontStyle glyphStyle = new FontStyle();
        private long numId = -1;
        private long lvl = -1;

        // 默认字体样式
        private FontStyle defaultTextStyle;

        public ParagraphBuilder() {

        }

        public ParagraphBuilder align(String align){
            String alignLower = align.toLowerCase();
            if("left".equals(alignLower)){
                this.align = STJc.LEFT;
            }else if("center".equals(alignLower)){
                this.align = STJc.CENTER;
            }else if("right".equals(alignLower)){
                this.align = STJc.RIGHT;
            }else {
                this.align = STJc.BOTH;
            }
            return this;
        }

        public ParagraphBuilder align(STJc.Enum align){
            this.align = align;
            return this;
        }

        public ParagraphBuilder identLeftChars(int value){
            this.indentLeftChars = value;
            return this;
        }

        public ParagraphBuilder indentRightChars(int value){
            this.indentRightChars = value;
            return this;
        }

        public ParagraphBuilder indentHangingChars(int value){
            this.indentHangingChars = value;
            return this;
        }

        public ParagraphBuilder indentFirstLineChars(int value){
            this.indentFirstLineChars = value;
            return this;
        }

        public ParagraphBuilder spacingBeforeLines(int value){
            this.spacingBeforeLines = value;
            return this;
        }

        public ParagraphBuilder spacingAfterLines(int value){
            this.spacingAfterLines = value;
            return this;
        }

        public ParagraphBuilder spacing(int value){
            this.spacing = value;
            return this;
        }

        public ParagraphBuilder spactingRule(String rule){
            String ruleLower = rule.toLowerCase();
            if("at_least".equals(ruleLower)){
                this.spacingRule = STLineSpacingRule.AT_LEAST;
            } else {
                this.spacingRule = STLineSpacingRule.AUTO;
            }
            return this;
        }

        public ParagraphBuilder spactingRule(STLineSpacingRule.Enum rule){
            this.spacingRule = rule;
            return this;
        }

        public ParagraphBuilder glyhStyle(FontStyle style){
            this.glyphStyle = style;
            return this;
        }

        public ParagraphBuilder numId(Long id){
            this.numId = id;
            return this;
        }

        public ParagraphBuilder lvl(long lvl){
            this.lvl = lvl;
            return this;
        }

        public ParagraphBuilder defaultStyle(FontStyle style){
            this.defaultTextStyle = style;
            return this;
        }

        public ParagraphBuilder spacingBefore(int spacingBefore){
            this.spacingBefore = spacingBefore;
            return this;
        }

        public ParagraphBuilder spacingAfter(int spacingAfter){
            this.spacingAfter = spacingAfter;
            return this;
        }

        public ParagraphBuilder indentHanging(int indentHanging){
            this.indentHanging = indentHanging;
            return this;
        }

        public ParagraphBuilder indentFirstLine(int indentFirstLine){
            this.indentFirstLine = indentFirstLine;
            return this;
        }

        public ParagraphStyle build(){
            return new ParagraphStyle(this);
        }
    }
}

```



# 段落样式
## 设置段落缩进
左侧缩进、右侧缩进以及特殊缩进方式或设置值
```java
    /**
     * 设置段落缩进 - 左侧缩进、右侧缩进以及特殊缩进方式及其值
     * @return
     */
    public static CTInd confIndBase(ParagraphStyle pStyle){
        CTInd ctInd = CTInd.Factory.newInstance();

        // 判断是什么缩进：首行/悬挂
        if(pStyle.getIndentFirstLine() > 0){
            // 首先判断是是否为首行缩进
            confTextFirstIndent(ctInd,pStyle.getIndentFirstLine(),pStyle.getIndentFirstLineChars());
        }else if(pStyle.getIndentHanging() > 0){
            // 判断是否为悬挂缩进
            confTextHanging(ctInd,pStyle.getIndentHanging(),pStyle.getIndentHangingChars());

        }

        // 设置左右测缩进
        // 左侧缩进
        ctInd.setLeftChars(BigInteger.valueOf(pStyle.getIndentLeftChars()));
        // 右侧缩进
        ctInd.setRightChars(BigInteger.valueOf(pStyle.getIndentRightChars()));

        return ctInd;
    }
```


```java
    /**
     * 设置首行缩进-特殊、缩进值
     */
    public static void confTextFirstIndent(CTInd ctInd,int iFL,int iFLC){
        ctInd.setFirstLine(BigInteger.valueOf(iFL));
        ctInd.setFirstLineChars(BigInteger.valueOf(iFLC));
    }
```

设置段落缩进 - 左侧缩进、右侧缩进以及特殊缩进方式及其值
```java
    /**
     * 设置段落缩进 - 左侧缩进、右侧缩进以及特殊缩进方式及其值
     */
    public static CTInd confIndBase(ParagraphStyle pStyle){
        CTInd ctInd = CTInd.Factory.newInstance();

        // 判断是什么缩进：首行/悬挂
        if(pStyle.getIndentFirstLine() > 0){
            // 首先判断是是否为首行缩进
            confTextFirstIndent(ctInd,pStyle.getIndentFirstLine(),pStyle.getIndentFirstLineChars());
        }else if(pStyle.getIndentHanging() > 0){
            // 判断是否为悬挂缩进
            confTextHanging(ctInd,pStyle.getIndentHanging(),pStyle.getIndentHangingChars());

        }

        // 设置左右测缩进
        // 左侧缩进
        ctInd.setLeftChars(BigInteger.valueOf(pStyle.getIndentLeftChars()));
        // 右侧缩进
        ctInd.setRightChars(BigInteger.valueOf(pStyle.getIndentRightChars()));

        return ctInd;
    }
```


```java
    /**
     * 设置首行缩进-特殊、缩进值
     */
    public static void confTextFirstIndent(CTInd ctInd,int iFL,int iFLC){
        ctInd.setFirstLine(BigInteger.valueOf(iFL));
        ctInd.setFirstLineChars(BigInteger.valueOf(iFLC));
    }
```


## 行间距
最小行间距
```java
    /**
     * 最小值行距
     * 设置段落行间距-段前、段后
     */
    public static CTSpacing confSpacingLeast(ParagraphStyle pStyle){
        CTSpacing ctSpacing = CTSpacing.Factory.newInstance();

        ctSpacing.setLine(BigInteger.valueOf(0)); // 最小值行距只需要设置一个line值
        ctSpacing.setLineRule(pStyle.getSpacingRule());
        return ctSpacing;
    }
```

单倍行距
```java
    /**
     * 单倍行距
     * 设置段落行间距-段前、段后
     */
    public static CTSpacing confSpacingSingle(ParagraphStyle pStyle){
        CTSpacing ctSpacing = CTSpacing.Factory.newInstance();

        ctSpacing.setLine(BigInteger.valueOf(240)); // 设置此条就会让行距规则变为 单倍行距

        ctSpacing.setLineRule(pStyle.getSpacingRule());
        return ctSpacing;
    }
```


1.5 倍行间距
```java
    /**
     * 设置段落行间距-段前、段后
     */
    public static CTSpacing confSpacingBase(ParagraphStyle pStyle){
        CTSpacing ctSpacing = CTSpacing.Factory.newInstance();
        ctSpacing.setBefore(BigInteger.valueOf(pStyle.getSpacingBefore()));
        ctSpacing.setAfter(BigInteger.valueOf(pStyle.getSpacingAfter()));
        ctSpacing.setBeforeLines(BigInteger.valueOf(pStyle.getSpacingBeforeLines()));
        ctSpacing.setAfterLines(BigInteger.valueOf(pStyle.getSpacingAfterLines()));
        ctSpacing.setLine(BigInteger.valueOf(360)); 
        ctSpacing.setLine(BigInteger.valueOf(pStyle.getSpacing()));
        ctSpacing.setLineRule(pStyle.getSpacingRule());
        return ctSpacing;
    }
```




# 字体样式
## 统一设置字体样式。修改各种编码
```java
    /**
     * 设置字体样式
     */
    public static void confFontFamily(CTFonts ctFonts, String fontFamily){
        ctFonts.setHint(STHint.EAST_ASIA);
        ctFonts.setEastAsia(fontFamily);
        ctFonts.setHAnsi(fontFamily);
        ctFonts.setAscii(fontFamily);
    }
```





# 清理所有

## 1.清理加粗

```java
    /**
    * 加粗清除器
    */
    public static void unsetBold(CTRPr ctrPr){
        if(ctrPr.isSetB()){
            ctrPr.unsetB();
        }

        if(ctrPr.isSetBdr()){
            ctrPr.unsetBdr();
        }

        if(ctrPr.isSetBCs()){
            ctrPr.unsetBCs();
        }
    }
```


## 2. 清理颜色
```java
    /**
     * 颜色清理
     */
    public static void unsetColor(CTRPr ctrPr){
        if(ctrPr.isSetColor()){
            ctrPr.unsetColor();
        }
    }
```


## 3. 斜体清理
```java
    /**
     * 斜体处理器
     */
    public static void unsetIn(CTRPr ctrPr){
        if(ctrPr.isSetI()){
            ctrPr.unsetI();
        }

        if(ctrPr.isSetICs()){
            ctrPr.unsetICs();
        }

        if(ctrPr.isSetImprint()){
            ctrPr.unsetImprint();
        }
    }
```


# 使用示例
示例项目待补，如果有时间或许后面也会补充个文字教程，顺便升级一下本项目的 poi 版本。



