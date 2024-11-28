<template>
  <div id="app" style="position: relative; height: 100vh;">
    <!-- 用于显示Excel内容的容器 -->
    <div id="luckysheet" style="padding: 0px; position: absolute; width: 100%; left: 0px; top: 10px; bottom: 10px"/>
  </div>
</template>

<script>
import LuckyExcel from 'luckyexcel';  // 引入LuckyExcel库
import * as XLSX from 'xlsx';  // 引入XLSX库用于解析Excel文件

export default {
  name: 'App',
  components: {},
  mounted() {
    // 组件挂载时加载Excel文件
    this.loadExcel();
  },
  data() {
    return {};
  },
  methods: {
    /**
     * 加载Excel文件
     */
    loadExcel() {
      fetch('/doc.xlsx')
          .then(response => response.blob())  // 获取Excel文件的Blob数据
          .then(blob => {
            const file = new File([blob], 'doc.xlsx', {type: blob.type});
            this.uploadExcel([file]);  // 上传Excel文件
          })
          .catch(error => {
            console.error('文件加载失败:', error);  // 错误处理
          });
    },

    /**
     * 解析并上传Excel文件
     * @param {Array} files 上传的文件数组
     */
    uploadExcel(files) {
      const file = files[0];

      // 使用FileReader读取Excel文件
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target.result;
        const workbook = XLSX.read(data, {type: 'array'});

        // 存储每个sheet的超链接信息
        const hyperlinks = this.extractHyperlinks(workbook);

        // 将文件传给LuckyExcel并进行转换
        LuckyExcel.transformExcelToLucky(file, (exportJson) => {
          // 在导出的JSON中设置超链接
          exportJson.sheets.forEach(sheet => {
            const sheetName = sheet.name;
            if (hyperlinks[sheetName]) {
              sheet.hyperlink = [];
              hyperlinks[sheetName].forEach(link => {
                this.addHyperlink(sheet, link.row, link.col, link.linkType, link.address, link.tooltip);
              });
            }
          });

          // 创建LuckyExcel实例并显示内容
          this.createLuckyExcel(exportJson);
        });
      };

      reader.readAsArrayBuffer(file);  // 以ArrayBuffer格式读取文件
    },

    /**
     * 提取Excel文件中所有sheet的超链接信息
     * @param {Object} workbook XLSX解析后的工作簿对象
     * @returns {Object} 每个sheet的超链接数据
     */
    extractHyperlinks(workbook) {
      const hyperlinks = {};
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        hyperlinks[sheetName] = this.extractSheetHyperlinks(sheet);
      });
      return hyperlinks;
    },

    /**
     * 提取指定sheet中的超链接
     * @param {Object} sheet sheet对象
     * @returns {Array} 超链接信息数组
     */
    extractSheetHyperlinks(sheet) {
      const hyperlinks = [];
      for (let cell in sheet) {
        const cellObj = sheet[cell];

        // 如果该单元格包含超链接
        if (cellObj && cellObj.l) {
          let address = cellObj.l.Target;  // 获取超链接地址
          const position = XLSX.utils.decode_cell(cell);  // 获取单元格位置

          // 判断链接是外部链接还是内部链接
          let linkType = address.startsWith('http://') || address.startsWith('https://') || address.startsWith('ftp://')
              ? 'external'
              : 'internal';

          // 对内部链接进行处理（去除#符号）
          if (linkType === 'internal') {
            address = address.replace('#', '');
          }

          // 存储超链接信息
          hyperlinks.push({
            row: position.r,
            col: position.c,
            address: address,
            tooltip: cellObj.v,  // 显示的文本
            linkType: linkType,
          });
        }

        // 也处理HYPERLINK函数中的超链接
        if (cellObj && cellObj.f && typeof cellObj.f === 'string') {
          const match = cellObj.f.match(/^HYPERLINK\("([^"]+)",\s*"([^"]+)"\)$/);
          if (match) {
            let address = match[1];
            const tooltip = match[2];
            const position = XLSX.utils.decode_cell(cell);

            let linkType = address.startsWith('http://') || address.startsWith('https://') || address.startsWith('ftp://')
                ? 'external'
                : 'internal';

            if (linkType === 'internal') {
              address = address.replace('#', '');
            }

            hyperlinks.push({
              row: position.r,
              col: position.c,
              address: address,
              tooltip: tooltip,
              linkType: linkType,
            });
          }
        }
      }
      return hyperlinks;
    },

    /**
     * 使用LuckyExcel创建并显示Excel
     * @param {Object} exportJson LuckyExcel导出的JSON数据
     */
    createLuckyExcel(exportJson) {
      window.luckysheet.destroy();  // 销毁之前的LuckyExcel实例
      window.luckysheet.create({
        lang: 'zh',  // 设置语言为中文
        data: exportJson.sheets,
        title: exportJson.info.name,
        userInfo: exportJson.info.name.creator,
        container: 'luckysheet',
        showtoolbar: false, // 是否显示工具栏
        showinfobar: false, // 是否显示顶部信息栏
        showstatisticBar: true, // 是否显示底部计数栏
        sheetBottomConfig: true, // sheet页下方的添加行按钮和回到顶部按钮配置
        allowEdit: false, // 是否允许前台编辑
        enableAddRow: false, // 是否允许增加行
        enableAddCol: false, // 是否允许增加列
        sheetFormulaBar: false, // 是否显示公式栏
        enableAddBackTop: false, // 返回头部按钮
        showsheetbar: false, // 是否显示底部sheet页按钮
      // 自定义配置底部sheet页按钮
        showsheetbarConfig: {
          add: false,
          menu: false,
        },
      });
    },

    /**
     * 在指定sheet的单元格中添加超链接
     * @param {Object} sheet sheet对象
     * @param {Number} row 行号
     * @param {Number} col 列号
     * @param {String} type 链接类型（internal或external）
     * @param {String} address 超链接地址
     * @param {String} tooltip 超链接的提示文本
     */
    addHyperlink(sheet, row, col, type, address, tooltip = '') {
      if (!sheet.hyperlink) {
        sheet.hyperlink = {};  // 初始化超链接对象
      }

      // 添加超链接
      sheet.hyperlink[`${row}_${col}`] = {
        linkType: type,
        linkAddress: address,
        linkTooltip: tooltip,
      };
      console.log('添加超链接完成:', sheet.hyperlink);
    },
  }
}
</script>

<style>
html, body {
  height: 100%;
  margin: 0;
  padding: 0;
}

#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  height: 100%;
}
</style>
