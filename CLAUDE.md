# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个基于 Tauri 2 的桌面应用程序，用于转换 Excel 文件格式。应用将源 Excel 文件中的数据通过可配置的列映射转换为标准化的收件人/订单格式，支持重复项检测和合并功能。

## 技术栈

- **前端**: 原生 HTML/CSS/JavaScript（无框架）
- **后端**: Rust + Tauri 2
- **Excel 处理**:
  - 读取: `calamine` (0.26)
  - 写入: `rust_xlsxwriter` (0.79)

## 常用命令

```bash
# 开发模式运行（热重载）
npm run tauri:dev

# 构建生产版本
npm run tauri:build

# 直接调用 Tauri CLI
npm run tauri -- [命令]
```

## 架构说明

### 前端结构

- **index.html**: 单页应用入口，定义了表格预览和列映射配置面板
- **src/main.js**: 前端逻辑
  - 通过 Tauri API 调用后端命令 (`invoke`)
  - 使用 `@tauri-apps/plugin-dialog` 进行文件选择
  - 管理 UI 状态和数据渲染
- **src/style.css**: 样式文件

### 后端结构（Rust）

核心模块位于 `src-tauri/src/`：

1. **main.rs**: 应用入口，调用 lib.rs 的 `run()` 函数

2. **lib.rs**: Tauri 应用核心
   - 定义 `AppState`：使用 `Mutex` 管理共享状态（转换后的数据和源文件路径）
   - 注册 Tauri 命令处理器：
     - `read_columns`: 读取 Excel 第一行作为列信息
     - `convert_with_mapping`: 根据用户配置的列映射转换文件
     - `convert_file`: 使用硬编码映射转换（向后兼容）
     - `export_file`: 导出转换后的数据为 Excel
     - `merge_duplicates`: 合并重复收件人的订单

3. **excel.rs**: Excel 处理核心逻辑
   - **数据模型**:
     - `ColumnInfo`: 列信息（索引、列号、标题）
     - `ColumnMapping`: 列映射配置（源列索引数组 + 操作类型）
     - `ConvertedRow`: 转换后的标准化行数据
     - `ConversionResult`: 转换结果（包含重复项统计）

   - **关键函数**:
     - `read_columns()`: 读取 Excel 第一行，返回所有列的信息
     - `read_and_convert_with_mapping()`: 主转换函数
       - 根据 `mappings: HashMap<String, ColumnMapping>` 映射源列到目标字段
       - 支持字符串拼接和数值计算（加减乘除）
       - 自动排序和标记重复项（按姓名+手机+地址）
     - `merge_duplicates()`: 合并重复收件人
       - 商品名称、规格用 "；" 拼接并去重
       - 数量求和（多商品时各为1）
       - 备注拼接并去重
     - `write_output()`: 写入标准化的 Excel 文件

### 数据流

1. 用户选择源 Excel 文件
2. 调用 `read_columns` 获取所有列信息
3. 前端显示列映射配置面板，用户为每个目标字段选择源列和操作
4. 调用 `convert_with_mapping` 执行转换
5. 转换结果保存在 `AppState` 中，前端显示预览表格
6. 用户可选择合并重复项或直接导出

### 重复项检测逻辑

- **检测依据**: (收件人姓名, 手机号, 收货地址) 三元组
- **标记**: 重复项会被分配 `group_id`（从1开始），用于前端高亮显示
- **排序**: 数据按三元组排序，使重复项相邻

### 列映射系统

- 每个目标字段可以映射到**多个源列**
- 支持的操作：
  - 字符串字段: `concat`（拼接）
  - 数值字段: `add`（加）、`subtract`（减）、`multiply`（乘）、`divide`（除）
- 前端使用多选下拉框让用户选择源列

## Tauri 配置

- **配置文件**: `src-tauri/tauri.conf.json`
- **前端资源目录**: `dist/`（注意：构建时需要先生成此目录）
- **应用标识**: `com.xlsconvert.tool`
- **窗口尺寸**: 1000x700，可调整大小
- **全局 Tauri**: 启用 (`withGlobalTauri: true`)

## 开发注意事项

### 前端与后端通信

前端通过 `window.__TAURI__.core.invoke` 调用后端命令：
```javascript
await invoke("convert_with_mapping", {
  path: filePath,
  mappings: { /* ... */ }
});
```

参数名必须与 Rust 函数签名中的参数名匹配（使用 camelCase）。

### 状态管理

AppState 使用 `Mutex` 保护共享数据，所有修改都需要获取锁。如果遇到死锁，检查是否有未释放的锁。

### Excel 数据类型处理

`cell_to_string()` 函数处理 `calamine::Data` 类型：
- 浮点数会移除不必要的小数点（如手机号）
- 空单元格返回空字符串
- 日期时间类型会被格式化为字符串

### 构建前提

由于 `tauri.conf.json` 中 `frontendDist` 指向 `../dist`，构建前需要确保该目录存在。当前项目前端资源直接在根目录，可能需要调整构建流程或配置。
