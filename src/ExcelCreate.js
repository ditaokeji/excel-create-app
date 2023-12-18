import React, {useState} from 'react'
import {utils, writeFile} from 'xlsx-js-style'
import moment from 'moment'
import {Button} from 'antd'
import {sequences, titles, columns, datas, infos} from './data'
// import Toast from "../Toast";

const ExcelCreate = (props) => {
  const {writeFinish, queryData, headerColumns, dataList, fileName = '模板', sheetName} = props

  const [loading, setLoading] = useState(false)

  // 导出后回调函数
  const writeCallback = (e, failure) => {
    if (writeFinish && typeof (writeFinish) == 'function')
      writeFinish(e, failure)
  }

  const exportFile = async() => {
    const time = moment(new Date()).format('YYYY-MM-DD HHmmss')
    let file = `${fileName}${time}.xlsx`

    setLoading(true)
    try{
      const sequencesArr = sequences.map(value =>{
        return [
          {
            v: value,
            t: 's',
            s: {
              font: {
                sz: 15, //设置标题的字号
                bold: true, //设置标题是否加粗
                name: '宋体',
              },
              //设置标题水平竖直方向居中，并自动换行展示
              alignment: {
                horizontal: 'center',
                vertical: 'center',
                wrapText: true,
              },
              border: {
                top: { style: 'thin', color: { rgb: '000000' } },
                bottom: { style: 'thin', color: { rgb: '000000' } },
                left: { style: 'thin', color: { rgb: '000000' } },
                right: { style: 'thin', color: { rgb: '000000' } },
              },
            },
          },
        ]
      })

      const titlesArr = titles.map(record => {
        let titles = []
        record.forEach(item => {
          titles.push({
            v: item.text,
            t: 's',
            s: {
              border: {
                top: { style: 'thin', color: { rgb: '000000' } },
                bottom: { style: 'thin', color: { rgb: '000000' } },
                left: { style: 'thin', color: { rgb: '000000' } },
                right: { style: 'thin', color: { rgb: '000000' } },
              },
            }
          })
          for(let i=0; i<item.span-1;i++){
            titles.push(null)
          }
        })
        return titles
      })

      const columnsArr = columns.map((e) => {
        return {
          v: e.text,
          t: 's',
          s: {
            font: {
              bold: true, //设置标题是否加粗
              name: '宋体',
            },
            //设置标题水平竖直方向居中，并自动换行展示
            alignment: {
              horizontal: 'center',
              vertical: 'center',
              wrapText: true,
            },
            border: {
              top: { style: 'thin', color: { rgb: '000000' } },
              bottom: { style: 'thin', color: { rgb: '000000' } },
              left: { style: 'thin', color: { rgb: '000000' } },
              right: { style: 'thin', color: { rgb: '000000' } },
            },
          },
        }
      })

      const dataArr = () => {
        const items= []
        datas.map((e, index) => {
          const item= []
          columns.map((ele, idx) => {
            item.push({
              v: e[ele.dataIndex],
              t: ele.summaryType=='sum'?'n':'s',
              s: {
                border: {
                  top: { style: 'thin', color: { rgb: '000000' } },
                  bottom: { style: 'thin', color: { rgb: '000000' } },
                  left: { style: 'thin', color: { rgb: '000000' } },
                  right: { style: 'thin', color: { rgb: '000000' } },
                },
              },
            })
          })
          items.push(item)
        })
        return items
      }

      let sumRow = 0
      const sumArr = Array.apply(null,{length: columns.length}).map((e,index) => {
        columns[index].summaryType === 'sum' && (sumRow=1)
        return {
          v: index==0?'合计':'',
          t: 's',
          s: {
            border: {
              top: { style: 'thin', color: { rgb: '000000' } },
              bottom: { style: 'thin', color: { rgb: '000000' } },
              left: { style: 'thin', color: { rgb: '000000' } },
              right: { style: 'thin', color: { rgb: '000000' } },
            },
          },
        }
      })

      const infosArr = infos.map(record => {
        let infos = []
        record.forEach(item => {
          infos.push({
            v: item.text,
            t: 's',
            s: {
              fill: {
                fgColor: { rgb: 'ffffff' },
              },
            }
          })
          for(let i=0; i<item.span-1;i++){
            infos.push(null)
          }
        })
        return infos
      })

      const sheetData = [...sequencesArr,...titlesArr,[...columnsArr],...dataArr()]
      sumRow && sheetData.push([...sumArr])
      sheetData.push(...infosArr)
      const ws = utils.json_to_sheet(sheetData, {skipHeader: true}) //

      // 合并单元格
      if (!ws['!merges']) ws['!merges'] = []
      sequences.forEach((value, index) => {
        ws['!merges']?.push(utils.decode_range(`A${index+1}:${String.fromCharCode('A'.charCodeAt(0)+columns.length-1)}${index+1}`))
      })

      let titlesMergesNum = 0
      titles.forEach((record, index) => {
        titlesMergesNum+=record.length
        const lastRowIndex = sequences.length
        let colSpan = 0
        record.forEach((item, idx) => {
          const startColIndex = String.fromCharCode('A'.charCodeAt(0)+colSpan)
          colSpan+=item.span
          const endColIndex = String.fromCharCode('A'.charCodeAt(0)+colSpan-1)
          ws['!merges']?.push(utils.decode_range(`${startColIndex}${lastRowIndex+index+1}:${endColIndex}${lastRowIndex+index+1}`))
        })
      })

      columns.forEach((item,index) => {
        if(item.summaryType === 'sum'){
          const colIndex = String.fromCharCode('A'.charCodeAt(0)+index)
          const columnsRow = 1
          const rowIndex = sequences.length+titles.length+datas.length+columnsRow+sumRow
          utils.sheet_set_array_formula(ws, `${colIndex}${rowIndex}`, `SUM(${colIndex}${rowIndex-datas.length}:${colIndex}${rowIndex-1})`, true)
        }
      })

      infos.forEach((record, index) => {
        const columnsRow = 1
        const lastRowIndex = sequences.length+titles.length+datas.length+columnsRow+sumRow
        let colSpan = 0
        record.forEach((item, idx) => {
          const startColIndex = String.fromCharCode('A'.charCodeAt(0)+colSpan)
          colSpan+=item.span
          const endColIndex = String.fromCharCode('A'.charCodeAt(0)+colSpan-1)
          ws['!merges']?.push(utils.decode_range(`${startColIndex}${lastRowIndex+index+1}:${endColIndex}${lastRowIndex+index+1}`))
        })
      })

      //给合并行列赋值样式
      const addRangeBorder = (range, ws) => {
        let cols = Array.apply(null, { length: 26 }).map((value, index) => String.fromCharCode('A'.charCodeAt(0)+index))
        let style = {
          v: '',
          t: 's',
          s: {
            border: {
              top: { style: 'thin', color: { rgb: '000000' } },
              bottom: { style: 'thin', color: { rgb: '000000' } },
              left: { style: 'thin', color: { rgb: '000000' } },
              right: { style: 'thin', color: { rgb: '000000' } },
            }
          }
        }
        range.forEach((item) => {
          let startRowNumber = Number(item.s.c)
          let endRowNumber = Number(item.e.c)
          for(let i = startRowNumber+1; i<=endRowNumber;i++) {
            ws[cols[i]+(Number(item.e.r)+1)]=style
          }
        })
        return ws
      }
      // 表格下面的说明-合并行列不用给边框
      const rangeBorders = ws['!merges'].slice(0, titlesMergesNum+sequences.length)
      addRangeBorder(rangeBorders,ws)

      // 设置列宽
      // cols 为一个对象数组，依次表示每一列的宽度
      if (!ws['!cols']) ws['!cols'] = []
      ws['!cols'] = columns.map(item => {
        if(item.width) {
          return { wpx: item.width }
        }
        return { wpx: 40 }
      })

      // 设置行高
      // rows 为一个对象数组，依次表示每一行的高度
      if (!ws['!rows']) ws['!rows'] = []
      console.log(sequences.length)
      console.log(titles.length)
      console.log(columns.length)
      console.log(datas.length)
      console.log(infos.length)
      ws['!rows'] = [
        ...Array.apply(null, { length: sequences.length }).map(() => {
          return { hpx: 40 }
        }),
        ...Array.apply(null, { length: titles.length }).map(() => {
          return { hpx: 20 }
        }),
        ...Array.apply(null, { length: 1 }).map(() => {
          return { hpx: 25 }
        }),
        ...Array.apply(null, { length: datas.length }).map(() => {
            return { hpx: 18 }
        }),
        ...Array.apply(null, { length: infos.length }).map(() => {
          return { hpx: 18 }
      }),
      ]

      const wb = utils.book_new()
      utils.book_append_sheet(wb, ws, sheetName||fileName)

      writeFile(wb, file) //
      writeCallback([[...columnsArr],...dataArr()])
      setLoading(false)
    }catch(err) {
      writeCallback(null, () => err)
      setLoading(false)
    }
  }

  return <>
    <Button loading={loading} onClick={exportFile}>订单导入</Button>
  </>
}

export default ExcelCreate
