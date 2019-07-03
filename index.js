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
            // 暂存多个邮箱的行
            for (let i = 0; i < sheet.data.length; i++) {
                if(sheet.data[i][0] == undefined){
                    // 行尾
                    break;
                }
                if(sheet.data[i][14] == undefined){
                    // 跳过 没有邮箱的列
                    continue;
                }
                var emails = sheet.data[i][14].trim();
                if(emails[emails.indexOf(";",emails.length - 1)] == ";"){
                    emails = emails.slice(0,emails.length - 2); 
                }
                emails = emails.split(";");
                for (let n = 0; n < emails.length; n++) {
                    if(emails[n] != '' && emails[n] != '...'){
                        tempSheet.push([emails[n].trim(),sheet.data[i][0]]);
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
