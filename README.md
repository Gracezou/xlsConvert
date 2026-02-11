# Excel 文件转换工具

一个基于 Tauri 开发的桌面应用程序，用于将自定义格式的 Excel 文件转换为标准化的收件人/订单格式。支持灵活的列映射配置、重复项检测和智能合并功能。

## ✨ 功能特性

- 📊 **灵活的列映射** - 可视化配置源 Excel 列到目标字段的映射关系
- 🔢 **多列操作** - 支持多个源列拼接或数值计算（加减乘除）
- 🔍 **重复项检测** - 自动识别相同收件人的多个订单
- 🔀 **智能合并** - 一键合并重复收件人的订单信息
- 👁️ **实时预览** - 转换后数据即时显示在表格中
- 💾 **快速导出** - 生成符合标准格式的 Excel 文件

## 🚀 快速开始

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run tauri:dev
```

应用将以开发模式启动，支持热重载。

### 构建应用

```bash
npm run tauri:build
```

构建产物位于 `src-tauri/target/release/bundle/` 目录。

## 📖 使用说明

1. **选择源文件** - 点击"选择源文件"按钮，选择要转换的 Excel 文件（支持 .xlsx 和 .xls 格式）

2. **配置列映射** - 在弹出的配置面板中：
   - 为每个目标字段选择对应的源列（支持多选）
   - 选择操作类型：
     - 字符串字段：拼接
     - 数值字段：加/减/乘/除

3. **应用转换** - 点击"应用配置并转换"，数据将显示在预览表格中

4. **合并重复项（可选）** - 如果检测到重复收件人，可点击"合并重复项"按钮进行智能合并

5. **导出文件** - 点击"导出文件"按钮，选择保存位置即可生成标准格式的 Excel 文件

## 🎯 输出格式

转换后的 Excel 文件包含以下标准字段：

| 字段 | 说明 | 是否必填 |
|------|------|----------|
| 收件人姓名 | 收件人的姓名 | 是 |
| 收件人手机号 | 收件人的联系电话 | 是 |
| 收货地址 | 完整的收货地址 | 是 |
| 商品名称 | 商品名称，多商品用"；"隔开 | 是 |
| 商品规格 | 商品规格，多商品用"；"隔开 | 否 |
| 商品数量 | 商品数量，多商品用"；"隔开 | 是 |
| 备注 | 额外备注信息 | 否 |

## 🛠️ 技术栈

- **前端**: HTML / CSS / JavaScript (原生)
- **后端**: Rust + Tauri 2
- **Excel 读取**: [calamine](https://github.com/tafia/calamine) 0.26
- **Excel 写入**: [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) 0.79

## 📋 系统要求

- Node.js 16+
- Rust 1.70+
- 操作系统：Windows / macOS / Linux

## 🔧 开发指南

详细的开发文档请参考 [CLAUDE.md](./CLAUDE.md)

### 项目结构

```
xlsConvert/
├── src/                  # 前端源代码
│   ├── main.js          # 主逻辑
│   └── style.css        # 样式文件
├── src-tauri/           # Tauri 后端
│   ├── src/
│   │   ├── main.rs      # 入口文件
│   │   ├── lib.rs       # 核心逻辑
│   │   └── excel.rs     # Excel 处理模块
│   └── tauri.conf.json  # Tauri 配置
├── index.html           # 应用入口页面
└── package.json         # 项目配置

```

### 添加新功能

1. 如需修改转换逻辑，编辑 `src-tauri/src/excel.rs`
2. 如需调整 UI，编辑 `src/main.js` 和 `src/style.css`
3. 如需添加新的 Tauri 命令，在 `src-tauri/src/lib.rs` 中定义并注册

## 📝 许可证

MIT

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！
