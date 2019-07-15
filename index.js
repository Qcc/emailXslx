const xlsx = require('node-xlsx')
const fs=require('fs');
//readdir为读取该文件夹下的文件
fs.readdir('./input', function(err,files){
    files.forEach((file) => {
        let path = `${__dirname}/input/${file}`;
        console.log(path);
        //表格解析
        let sheetList = xlsx.parse(path);
        //对数据进行处理
        sheetList.forEach((sheet) => {
            // 删除列名称字段
            sheet.data.splice(0,1);
            tempSheet = [];
            // 暂存多个电话的行
            for (let i = 0; i < sheet.data.length; i++) {
                if(sheet.data[i][0] == undefined){
                    // 行尾
                    break;
                }
                if(sheet.data[i][10] == undefined){
                    // 跳过 没有电话的列
                    continue;
                }
                if(sheet.data[i][11] != undefined){
                    sheet.data[i][10] = sheet.data[i][10] + ";" + sheet.data[i][11]; 
                }
                var phones = sheet.data[i][10].trim();
                if(phones[phones.indexOf(";",phones.length - 1)] == ";"){
                    phones = phones.slice(0,phones.length - 1); 
                }
                phones = phones.split(";");
                var contact = sheet.data[i][0].slice(0,sheet.data[i][0].indexOf("有限公司")) +"-"+ sheet.data[i][1];
                for (let n = 0; n < phones.length; n++) {
                    if(phones[n] != '' && phones[n].length == 11 && phones[n].indexOf("-") == -1 && phones[n] != '...'){
                        tempSheet.push([contact,phones[n].trim()]);
                    };
                }
            }
            sheet.data = tempSheet;
        })
        //数据进行缓存
        let buffer = xlsx.build(sheetList);
        //将缓存的数据写入到相应的Excel文件下
        fs.writeFile(path.replace(/input/, 'output').replace(/\./, '_修改版.'), buffer, function(err){ 
            if (err) {
                console.log(err);
                return ;
            }
        });
    })
});
