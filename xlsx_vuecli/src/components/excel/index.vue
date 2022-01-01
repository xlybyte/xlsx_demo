<template>
  <div class="dashboard-container">
    <div class="app-container">
      <!-- Excel导入组件 -->
      <upload-excel-component :on-success="handleSuccess" :before-upload="beforeUpload" />
      <!-- 导出按钮 -->
      <el-button type="success" size="small" @click="exportExcelFn">导出excel</el-button>
      <!-- 表格显示数据 -->
      <el-table border :data="dataArr">
          <el-table-column label="序号" type="index" />
          <el-table-column label="姓名" prop="username" />
          <el-table-column label="手机号" prop="mobile" width="120" />
          <el-table-column label="工号" prop="workNumber" sortable :sort-method="sortWorkNumberFn" />
          <el-table-column label="聘用形式" prop="formOfEmployment" />
          <el-table-column label="部门" prop="departmentName" />
          <el-table-column label="入职时间" prop="timeOfEntry"/>
      </el-table>
    </div>
  </div>
</template>

<script>
// import { formatExcelDate } from '@/utils/index'
import UploadExcelComponent from '@/components/UploadExcel/index'
export default {
  components: { UploadExcelComponent },
  data () {
    return {
      // 导入的EXCEL数据
      tableData: [], // 数据
      tableHeader: [], // 头
      dataArr: [] // 转换为英文后的数据
    }
  },
  created () {
    console.log(import('@/vendor/Export2Excel'))
  },
  methods: {
    // 导入excel 之前事件
    beforeUpload (file) {
      const isLt1M = file.size / 1024 / 1024 < 1
      if (isLt1M) return true
      this.$message({
        message: '请不要上传大小超过1m的文件.',
        type: 'warning'
      })
      return false
    },
    // 导入excel成功事件
    async handleSuccess ({ results, header }) {
      this.tableData = results
      this.tableHeader = header
      console.log('header', header)
      console.log('results', results)
      // 转换格式
      const arr = this.transExcel(results)
      this.dataArr = arr
      console.log('转换格式', arr)
      // const res = await importEmployeeAPI(arr)
    },
    // excel数据转提交格式
    transExcel (results) {
      const userRelations = {
        入职日期: 'timeOfEntry',
        手机号: 'mobile',
        姓名: 'username',
        转正日期: 'correctionTime',
        工号: 'workNumber',
        部门: 'departmentName',
        聘用形式: 'formOfEmployment'
      }
      const arr = []
      results.forEach(item => {
        const obj = {}
        const contentKeys = Object.keys(item)
        contentKeys.forEach(k => {
          const key = userRelations[k]
          if (key) {
            // 如果时间格式为数字，则需要调用函数转换
            // if (key === 'timeOfEntry' || key === 'correctionTime') {
            //   item[k] = formatExcelDate(item[k], '-')
            // }
            obj[key] = item[k]
          }
        })
        arr.push(obj)
      })
      return arr
    },
    // 工号排序函数 - 自定义排序
    sortWorkNumberFn (a, b) {
      return a.workNumber - b.workNumber
    },
    // 导出excel
    exportExcelFn () {
      // 导出的列顺序和数组顺序一致，key为数据中的属性，name为导出后的列名
      const xlsHeader = [
        { key: 'id', name: '编号' },
        { key: 'username', name: '姓名' },
        { key: 'staffPhoto', name: '头像地址' },
        { key: 'mobile', name: '手机号' },
        { key: 'workNumber', name: '工号' },
        { key: 'formOfEmployment', name: '聘用形式' },
        { key: 'departmentName', name: '部门' },
        { key: 'timeOfEntry', name: '入职日期' }
      ]
      // 调用方法导出
      import('@/vendor/Export2Excel').then(excel => {
        const list = this.dataArr // 要导出的数组对象
        const tHeader = xlsHeader.map(obj => obj.name) // 遍历出表头
        const data = list.map((obj, index) => {
          // obj 为每一行数据对象
          return xlsHeader.map(v => {
            // 自定义对每一列数据进行处理
            if (v.key === 'id') return index + 1
            // if (v.key === 'formOfEmployment') return this.formatEmployeeFn(obj[v.key])
            // if (v.key === 'timeOfEntry') return parseTime(obj[v.key], '{y}-{m}-{d}')
            return obj[v.key]
          })
        })
        console.dir(excel)
        // 开始导出
        excel.export_json_to_excel({
          header: tHeader, // 导出的表头，['id', '姓名']
          data, // 导出的数据，数组套数组格式。[['1', '张三'], ['2', '李四']]
          filename: 'xlsxxlsx', // 文件名
          autoWidth: true, // 是否自动列宽
          bookType: 'xlsx' // 格式
        })
      })
      // 然后浏览器会弹出下载excel文件
    }
  }
}
</script>

<style>

</style>
