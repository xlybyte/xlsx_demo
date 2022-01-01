/**
 * 把excel文件中的日期格式的内容转回成标准时间
 * https://blog.csdn.net/qq_15054679/article/details/107712966
 * @param {*} numb 待转换的数字
 * @param {*} format 分隔符
 * @returns 返回转换后的日期格式
 */
export function formatExcelDate (numb, format = '/') {
  const time = new Date((numb - 25567) * 24 * 3600000 - 5 * 60 * 1000 - 43 * 1000 - 24 * 3600000 - 8 * 3600000)
  time.setYear(time.getFullYear())
  const year = time.getFullYear() + ''
  const month = time.getMonth() + 1 + ''
  const date = time.getDate() + ''
  if (format && format.length === 1) {
    return year + format + month + format + date
  }
  return year + (month < 10 ? '0' + month : month) + (date < 10 ? '0' + date : date)
}
