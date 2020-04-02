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
            
            arr_data.push(line.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g));
            
        });
        rl.on('close', function() {
            
            var workbook = new ExcelJS.Workbook();
            var sheet = workbook.addWorksheet('Sheet1');
            arr_data.forEach((r)=>{
                for(k in r){
                    r[k]=r[k].replace(/['"]+/g, '');
                }
                sheet.addRow(r);
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