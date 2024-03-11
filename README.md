# fixWord
## 项目简介
fixWord是一个基于python开发的Word文档修复工具，能够自动修复Word文档中的常见错误，如拼写错误、格式错误等。
- [→访问主页（gitee）（推荐）](https://gitee.com/cxmStudio/fixWord)
- [→访问主页（github）](https://github.com/Sam-CXM/fixWord)
## 开发环境
- Python 3.10.7
- python-docx 1.1.0
- pyinstaller 5.6.1
## 项目特点
- 支持多种常见错误修复，如拼写错误、格式错误等。
- 支持自定义错误类型和修复规则。
- 支持单文件和批量修复多个Word文档。
- 支持自定义输出结果格式。
## 运行环境
| 系统 | 内存 | 磁盘 | 备注 |
| ---- | ---- | ---- | ---- |
| Windows10及以上版本 | 至少2GB | 至少25MB | / |
## 使用说明
1. [下载地址1（推荐）](https://gitee.com/cxmStudio/fixWord/releases/download/v1.2/fixWord_v1.2.zip) [下载地址2](https://github.com/Sam-CXM/fixWord/releases/download/v1.2/fixWord_v1.2.zip)
2. 将安装包解压到本地。
3. 运行 `fixWord_v1.2.exe` 文件，输入要修复的Word**文档路径**或含有文档的**文件夹路径**。
4. 功能选择（输入 `Y` 或 `y` 确定）。
5. 按回车键等待修复完成。
## 功能说明
- **开发作者**：晨小明
- **开发日期**：2024/01/04
- **开发版本**：v1.2_release
- **修改日期**：2024/03/11
- **主要功能**：
    - 一、支持**单文件处理**或**批量文档处理**，输入文件路径或文件夹路径，自动判断。
    - 二、**读取.docx文件并设置格式**
        1. 页边距：
            - 上3.7cm，
            - 下3.5cm，
            - 左2.8cm，
            - 右2.6cm
        2. 段落行距：
            - 标题：固定值33磅；
            - 正文：一般固定值28磅
        3. 字体，字号：
            - 标题：小二号方正小标宋简体，居中；
            - 一级标题：四号黑体；
            - 二级标题：四号楷体_GB2312；
            - 正文：四号仿宋_GB2312，两端对齐；
            - 数字&英文：四号TimesNewRoman字体
        4. 支持添加页码（可选）
            - 4号半角宋体阿拉伯数字，数字左右各加一条4号“一字线”，奇数页在右侧左空一字，偶数页在左侧左空一字
        5. 识别文档中的图片并输出（可选）
    - 三、**替换功能**
        1. 符号替换：
            将英文状态下的符号替换为中文状态下的相同符号，包含如下：
            - "`(`" --> "`（`"
            - "`)`" --> "`）`"
            - "`)、`" --> "`）`"
            - "`）、`" --> "`）`"
            - "`,`" --> "`，`"
            - "`:`" --> "`：`"
            - "`;`" --> "`；`"
            - "`?`" --> "`？`"
            - "[空格]" --> ""
        2. 其他格式
            数字后有顿号替换为点，如："1、" --> "1."
    - 四、输出文件名称含时间点，方便标记（可选）
## 更新日志：
  - 维护日期：**2024.3.12**
    - 【**修复**】解决了批量处理时选项需要重复输入的问题。
  - 维护日期：**2024.1.22**
    - 【**修复**】解决了含有图片的文档处理后图片被删除的问题。
  - 维护日期：**2024.1.21**
    - 【**新增**】可选项判断；
    - 【**新增**】处理完成后倒计时自动关闭；
    - 【**优化**】图片输出逻辑。
## 示例截图
### 基础功能
![功能示例后视图](/static/基础功能.png)
### 输出图片
![输出图片](/static/图片输出.png)
### 范文示例
![基本功能](/static/范文示例.png)
## 项目截图
![项目截图1](/static/项目截图1.png)
![项目截图2](/static/项目截图2.png)
![项目截图3](/static/项目截图3.png)
## 注意事项
- 本程序仅处理 `.docx` 类型的文件；
- 本程序暂不支持处理含有表格内容的文件；
- 含有图片的文档图片导出后可能会被压缩；
- 本程序无法处理图片格式，如果图片独立成段，本程序所用API识别到图片会被默认是空段落。为了防止图片删除，只能放弃处理空段落及图片格式；
- 为了处理效果，处理前请将全文`清除全部格式`，操作步骤：`全选`->`开始`->`样式`->`清除格式`；将文档中所有图片环绕文字改为`嵌入型`，操作步骤：`选中图片`->`图片格式`->`排列`->`环绕文字`->`嵌入型`；
- 本程序已开源，可免费使用。
## 源文档格式说明
### 标题格式
- 独立成段；
- 在文档的首行。
### 一级标题
- 独立成段；
- 以数字形式的汉字为段首字，其后加上中文形式的 `、` 号
    - 例如： `一、` `二、` ……
### 二级标题
- 独立成段；
- 以数字形式的汉字为段首字，其两边加上左右圆括号，中文或英文形式均可，程序会自动将英文格式的括号替换为中文形式。
    - `（` ` ）` （中文括号）
    -  `(` `)` （英文括号）
    - 例如： `（一）` `（二）` …… 或 `(一)` `(二)` ……
- 如果右括号后加 `、` 号，程序会自动删除。
### 数字/英文
- 数字为 `1` `2` `3` `4` `5` `6` `7` `8` `9` `0`；
- 英文为 `a` `b` `c` `d` `e` `f` `g` `h` `i` `j` `k` `l` `m` `n` `o` `p` `q` `r` `s` `t` `u` `v` `w` `x` `y` `z` `A` `B` `C` `D` `E` `F` `G` `H` `I` `J` `K` `L` `M` `N` `O` `P` `Q` `R` `S` `T` `U` `V` `W` `X` `Y` `Z` ；（26个英文大小写字母）
- 数字后如果有 `、` 号，程序会自动替换为 `.` 。
## 联系方式
- QQ: **3038693133**
- 邮箱：**3038693133@qq.com**
