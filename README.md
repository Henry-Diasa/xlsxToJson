<!--
 * @Author: diasa diasa@gate.me
 * @Date: 2025-05-30 14:17:56
 * @LastEditors: diasa diasa@gate.me
 * @LastEditTime: 2025-05-30 15:17:22
 * @FilePath: /crvToJson/README.md
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
-->
# Excel to JSON Translator

本工具可将当前目录下所有Excel（.xlsx）翻译表格批量转换为多语言JSON文件，并按文件名分目录输出。

## 功能特点

- 支持批量处理当前目录下所有`.xlsx`文件
- 每个Excel文件的翻译结果输出到`translations/文件名/`目录下
- 自动生成规范的扁平key（如`br.xx.yy`，总长度不超过16）
- 支持多语言，缺失内容自动补空字符串
- 只为有内容的列生成JSON文件

## 使用方法

1. 安装依赖：
```bash
npm install
```

2. 将所有需要转换的Excel文件（如`broker.xlsx`、`marketmaker.xlsx`等）放在项目根目录下。

3. 运行批量转换命令：
```bash
npm start
```

4. 程序会自动在`translations/文件名/`目录下生成各语言的JSON文件，例如：
```
translations/
  broker/
    zh.json
    en.json
  marketmaker/
    zh.json
    en.json
```

5. 也可单独处理某个文件：
```bash
npm start yourfile.xlsx
```

## Excel文件格式示例

| CN（简体） | EN                  | JA         |
|------------|---------------------|------------|
| 做市商项目 | Market Maker Program | ...        |
| 注册       | Register            | ...        |

## 生成的JSON示例

```json
{
  "br.ma.ma": "Market Maker Program",
  "br.re.xx": "Register"
}
```

## 说明
- key自动根据文件名前缀和内容生成，形式为`br.xx.yy`，总长度不超过16。
- 只为有内容的语言列生成JSON文件，内容全空的列不会生成。
- 支持多语言，缺失内容自动补空字符串。 