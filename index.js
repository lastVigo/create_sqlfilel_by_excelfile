let XLSX = require("xlsx");
let fs = require("fs");
let sd = require('silly-datetime');
const nanoid = require('nanoid');

// 生成当前时间字符串
var time = sd.format(new Date(), 'YYYYMMDDHHmm');
var buf = fs.readFileSync("./excelfiles/big.xlsx");
// 读取excel文档，存放在workbook中
var workbook = XLSX.read(buf, { type: 'buffer' });
// 读取excel工作簿数组
var sheetNames = workbook.SheetNames;
// 取其中第一个工作簿
var worksheet = workbook.Sheets[sheetNames[0]];
// 读取该工作簿的范围，返回一个由起止单元格名称组成的字符串（如'A1:B20'）
var rangeStr = worksheet['!ref'];
var patt1 = /\d+/g;
// 解析该字符串，分别得到起止行的行号
var lineArr = rangeStr.match(patt1);
let lineStart = lineArr[0];
var lineEnd = lineArr[1];

//根据单元格得到其中数据
let getCellValue = function(cell) {
    if (cell && cell.v) {
        let tempStr = cell.v;
        tempStr = tempStr.replace(/'/g, "");
        tempStr = tempStr.replace(/\n/g, " ");
        tempStr = tempStr.replace(/\r/g, " ");
        tempStr = tempStr.replace(/\\/g, " ");
        return tempStr;
    } else {
        return "";
    }
};
// sql前半部分
const insertHeadStr = `INSERT INTO crm_info (info_id, name,linker, link_tel, address,remark, age_id, email, data_resource,create_user,area) VALUES\r\n`;
// 根据数据生成sql后半部分
const getInsertContent = (data) => ` ('${data.info_id}', '${data.name}', '${data.linker}', '${data.link_tel}', '${data.address}','${data.remark}', '-1', '${data.email}', 2,'system_${data.time}','${data.area}')`;
//实际内容开始的行，表头占了两行
let index = 3;
// sql字符串buffer
let dataBuffer = "";
// 字符串buffer的最大记录数目
// 每100条生成一条sql语句
const BufferMaxLen = 100;
// 有效数据数量
let recordNum = 0;
//开始逐行遍历
for (let i = index, bufferLen = 1, dataBuffer = insertHeadStr; i <= lineEnd; i++) {
    let info_id = nanoid();
    // 读取name，如果没有值不生成SQL.
    let name = getCellValue(worksheet['A' + i]);
    if (name == null || name == "") {
        continue;
    }
    // 根据Excel结构，生成当前行的数据对象
    let data = {
        info_id: nanoid(),
        name,
        address: getCellValue(worksheet['B' + i]),
        link_tel: getCellValue(worksheet['C' + i]),
        email: getCellValue(worksheet['D' + i]),
        linker: getCellValue(worksheet['E' + i]),
        remark: getCellValue(worksheet['G' + i]),
        area: getCellValue(worksheet['I' + i]),
        time
    }
    dataBuffer += getInsertContent(data);
    bufferLen++;
    recordNum++;
    // buffer达到最大，追加写文件
    if (bufferLen == BufferMaxLen) {
        dataBuffer += `;\r\n`;
        fs.appendFileSync("./output/" + time + "_big.txt", dataBuffer);
        bufferLen = 0;
        dataBuffer = insertHeadStr;
    }
    // 到了最后一条，追加写文件
    else if (i == lineEnd) {
        dataBuffer += `;\r\n`;
        fs.appendFileSync("./output/" + time + "_big.txt", dataBuffer);
    }
    // 非以上情况
    else {
        dataBuffer += `,\r\n`;
    }

}
console.log(`sql文件生成完成,总共生成了${recordNum}条数据`);