const jaci = require("jaci");
const fs = require("fs");
const path = require("path");
const readline = require('readline')
const ExcelJS = require('exceljs');

const f_access = fs.constants.W_OK && fs.constants.F_OK;

var target_directory;
const jacsv =()=>{
    const loopFile=(idx,arr,cb)=>{
        if(idx>=arr.length){
            cb();
            return;
        }
        var sel = arr[idx];
        var fpath = target_directory +"/" + sel;
        var xlsxFile=path.basename(fpath).split(".")[0] + ".xlsx";
        var xlsxPath = target_directory +"/"+xlsxFile;
        var ext = path.extname(fpath).toLowerCase();
        var arr_data=[];
        if(ext !=".csv"){
            console.log(fpath,"Skipped, not csv");
            setTimeout(()=>{
                loopFile(idx+1,arr,cb);
            },1)
            return;
        };
        const rl = readline.createInterface({
            input: fs.createReadStream(fpath),
        });
        
        rl.on('line', function(line) {
            //line = line.replace(/['"]+/g, '');
            
            line = line.replace(/'/g,'')
            var test = line.split(",");
            if(line.indexOf('"')>=0){
                test = line.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g);
            }
            arr_data.push(test);
            
        });
        rl.on('close', function() {
            var workbook = new ExcelJS.Workbook();
            var sheet1 = workbook.addWorksheet('Sheet1');
            var sheet2= workbook.addWorksheet('Sheet2');
            arr_data.forEach((r)=>{

                var dir="";
                var z = JSON.parse(JSON.stringify(r));

                for(k in r){
                    r[k]=r[k].replace(/(\r\n|\n|\r)/gm,' ');
                    r[k]=r[k].replace(/['"]+/g, ' ');
                    r[k]=r[k].replace(/\s\s+/g, ' ');
                    z[k]=r[k];

                    if(k==1){
                        var t = r[k].indexOf(".00");
                        var x = r[k].substring(t).replace(".00","");
                        if(x.length>0){
                            r[k]=x;
                        }
                    }
                    if(k==3){
                        if(r[k].indexOf("DB")>=0){
                            dir="DB"
                        }
                        if(r[k].indexOf("CR")>=0){
                            dir="CR"
                        }
                        r[k] = r[k].replace(" DB","");
                        r[k] = r[k].replace(" CR","");
                    }
                }
                if(r.length >=3 && r[3]){
                    r[3] = r[3].replace(".00","");
                    r[4] = r[4].replace(".00","");
                    r[3] = r[3].replace(/,/g, '');
                    r[4] = r[4].replace(/,/g, '');

                    z[3] = z[3].replace(".00","");
                    z[4] = z[4].replace(".00","");
                    z[3] = z[3].replace(/,/g, '');
                    z[4] = z[4].replace(/,/g, '');

                }
                r.push(dir);
                sheet1.addRow(r);
                sheet2.addRow(z);
            });
            workbook.xlsx.writeFile(xlsxPath)
            .then(function() {
                setTimeout(()=>{
                    loopFile(idx+1,arr,cb);
                },1)
            });
            
        });
        
    }
    const run=()=>{
        jaci.string("Enter Directory Path : ",{required:false})
        .then((res)=>{
            target_directory = res;
            fs.access(target_directory, f_access, (err) => {
                if(err){
                    console.log("Directory is not accessible");
                    return;
                }
                fs.readdir(target_directory, function (err, files) {
                    //handling error
                    if (err) {
                        return console.log('Unable to scan directory: ' , err);
                    } 
                    loopFile(0,files,()=>{
                        console.log("Conversion success");
                        process.exit(-1);
                    })
                });
            });
        })
        .catch((e)=>{
            console.log("error",e);
            jaci.done();
        })
        return;
    }
    return {
        run:run
    }
}

module.exports = jacsv();