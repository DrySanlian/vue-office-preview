# vue-office-preview

一个支持 Vue 2、Vue 3 和原生 JavaScript 的办公文档预览组件库，支持 DOCX、XLSX/XLS、PDF 和 PPTX。

本仓库基于 [DrySanlian/vue-office-preview](https://github.com/DrySanlian/vue-office-preview) 开源项目进行二次开发和维护，感谢原项目提供的基础实现。

如果这个项目对你有帮助，欢迎在 GitHub 点一个 Star。你的 Star 会帮助更多人发现这个项目，也会鼓励持续维护和改进。

## 支持范围

| 文件类型 | Vue 组件 | 原生 JavaScript | 底层方案 |
| --- | --- | --- | --- |
| DOCX | `@vue-office/docx` | `@js-preview/docx` | `docx-preview` |
| XLSX / XLS | `@vue-office/excel` | `@js-preview/excel` | `exceljs` + `x-data-spreadsheet` |
| PDF | `@vue-office/pdf` | `@js-preview/pdf` | PDF.js |
| PPTX | `@vue-office/pptx` | 暂不提供 | 外部依赖 `pptx-preview` |

说明：当前重点支持 OOXML 格式的 `.pptx`，不承诺支持旧版二进制 `.ppt` 文件。

## 安装

按实际使用的文件类型安装对应包：

```shell
npm install @vue-office/docx vue-demi@0.14.6
npm install @vue-office/excel vue-demi@0.14.6
npm install @vue-office/pdf vue-demi@0.14.6
npm install @vue-office/pptx vue-demi@0.14.6
```

Vue 2.6 及以下版本还需要安装：

```shell
npm install @vue/composition-api
```

## Vue 使用

所有 Vue 组件都通过 `src` 接收文件内容，支持：

- 网络地址字符串
- `ArrayBuffer`
- `Blob`

以 PDF 为例：

```vue
<template>
  <vue-office-pdf
    :src="fileSource"
    style="height: 100vh"
    @rendered="handleRendered"
    @error="handleError"
  />
</template>

<script>
import VueOfficePdf from '@vue-office/pdf'

export default {
  components: { VueOfficePdf },
  data() {
    return {
      fileSource: 'https://example.com/document.pdf'
    }
  },
  methods: {
    handleRendered() {
      console.log('预览完成')
    },
    handleError(error) {
      console.error('预览失败', error)
    }
  }
}
</script>
```

通过接口获取二进制内容时，可以直接传入 `ArrayBuffer` 或 `Blob`：

```js
const response = await fetch('/api/file')
const fileSource = await response.arrayBuffer()
```

然后将 `fileSource` 绑定到组件的 `src` 属性即可。

## 当前维护内容

本 Fork 的维护重点是组件外层，不改变各文档底层解析库的职责：

- 统一处理 URL、`ArrayBuffer` 和 `Blob` 输入，并校验网络请求失败状态。
- 快速切换文件时使旧请求和旧渲染结果失效，避免旧内容覆盖新内容。
- 关闭或销毁预览时清理 PDF 渲染任务、Excel 全局事件监听、旧 DOM 和临时 `Object URL`。
- 修复重复打开、快速切换、关闭预览后鼠标短暂无法操作等生命周期问题。
- 保留大型 XLSX、TXT 等文件的预览能力，不在组件层设置统一的 10 MB 限制。
- Excel 默认使用 Worker 完成文件解析、格式转换和样式转换，主线程只负责表格数据加载和界面渲染。
- PPTX 只维护 `vue-pptx` 外层封装；`pptx-preview` 是外部依赖，本仓库不包含、不修改其内部实现。
- 构建脚本使用 Node.js 文件操作，支持 Windows 和 Unix 环境。

实际可预览大小仍受浏览器内存、文件复杂度、网络速度和底层解析库影响。大型文件建议在业务侧提供加载状态、错误提示和必要的超时处理。

### Excel Worker 说明

Excel Worker 默认启用，并支持快速切换和取消正在进行的解析。`transformData` 仍会在主线程执行；如果配置了 `beforeTransformData`，组件会自动回退到主线程，以保留对 ExcelJS 原始 `workbook` 对象的完整访问能力。

如果部署环境配置了严格的 CSP，需要允许当前站点加载同源 Worker。

## 本地开发

```shell
cd core
npm ci
npm run dev
```

构建单个组件：

```shell
npm run lib:vue-docx
npm run lib:vue-excel
npm run lib:vue-pdf
npm run lib:vue-pptx
npm run lib:js-docx
npm run lib:js-excel
npm run lib:js-pdf
```

构建产物位于对应包的 `lib/` 目录。Vue 组件会分别生成 Vue 2 和 Vue 3 版本。

## 源码结构

```text
core/
  packages/
    vue-docx/       Vue DOCX 组件
    vue-excel/      Vue Excel 组件
    vue-pdf/        Vue PDF 组件
    vue-pptx/       Vue PPTX 外层组件
    js-docx/        原生 JavaScript DOCX 预览
    js-excel/       原生 JavaScript Excel 预览
    js-pdf/         原生 JavaScript PDF 预览
  utils/            URL、下载和通用工具
  script/           构建辅助脚本
```

## 维护边界和许可证

本项目是在原开源项目基础上的二次开发。使用、修改和发布时，请同时遵守本仓库 LICENSE、原项目 LICENSE 以及各第三方依赖对应的许可证。

PPTX 的底层 `pptx-preview` 依赖不属于本仓库源码；本 Fork 只维护 `core/packages/vue-pptx` 的输入适配、生命周期、错误处理和资源清理。
