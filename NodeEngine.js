console.clear()
// manipulation node js File System
const fs= require('fs')
const path=require('path')
const dirPath=path.join(__dirname,'dataExcel')
// obtenir la liste des fichiers xlsx
var listFiles=fs.readdirSync(dirPath)
// console.log(listFiles);
// filtre pour les fichier de 09h01
function rapportTitreFiltre(fileList){
    let filtredList=[]
    for (let i = 0; i < fileList.length; i++) {
        const nombreHeure=parseInt(fileList[i].slice(22, 24))
        if ( nombreHeure <=10 && nombreHeure>=9) {
            filtredList.push(fileList[i])
        }
    }
    return filtredList
}

// lecture fichier excel
var  xlsx= require("xlsx")
const fileName='Rapport-PRDD-20220301-0901.xlsx'

 console.log(importExcelData(fileName, dirPath)); 

function importExcelData(nameOfFile, direrctoryfile){
    var wb=xlsx.readFile(dirPath+'/'+fileName,{cellDates:true})
    var nomsPages=wb.SheetNames
    // lecture de la feuil par defautl
    ws=wb.Sheets[nomsPages[0]]
    //convertir à json
    var data=xlsx.utils.sheet_to_json(ws)
    //enlèvement des lignes unitiles 
    data=data.splice(2,data.length)
    return data
}







// var newData=data.map((record)=>{
//     var test=new Date(record.__EMPTY_10)
//     record.Test=test.getFullYear()
//     return record
// })

// newData.splice(0,2)
// // console.log(newData);
// var newData2=newData

// var newWB= xlsx.utils.book_new();
// var newWS= xlsx.utils.json_to_sheet(newData)
// xlsx.utils.book_append_sheet(newWB,newWS,"NEW Data")
// xlsx.writeFile(newWB,'dataExcel/newdataFile2.xlsx')

// var newWS= xlsx.utils.json_to_sheet(newData2)
// xlsx.utils.book_append_sheet(wb,newWS,"NEW Data")
// xlsx.writeFile(wb,'dataExcel/newdataFile3.xlsx')