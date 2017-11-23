var fs = require('fs');
var xlsx = require('xlsx');
var cvcsv = require('csv');

exports = module.exports = XLSX_json;

function XLSX_json (config, callback, transform) {
  if(!config.input) {
    console.error("You miss a input file");
    process.exit(1);
  }

  var cv = new CV(config, callback, transform);

}

function CV(config, callback, transform) {
  var wb = this.load_xlsx(config.input)
  var ws = this.ws(config, wb);
  var csv = this.csv(ws)
  this.cvjson(csv, config, callback, transform)
}

CV.prototype.load_xlsx = function(input) {
  return xlsx.readFile(input);
}

CV.prototype.ws = function(config, wb) {
  var target_sheet = config.sheet;

  if (target_sheet == null)
    target_sheet = wb.SheetNames[0];

  ws = wb.Sheets[target_sheet];
  return ws;
}

CV.prototype.csv = function(ws) {
  return csv_file = xlsx.utils.make_csv(ws)
}

CV.prototype.createFile = function(output,record,callback){
  var stream = fs.createWriteStream(output, {flags : 'w'});
  stream.write(JSON.stringify(record));
  callback(null, record);
}

CV.prototype.cvjson = function(csv, config, callback, transform) {
  var record = [];
  var header = [];
  var self = this;

  cvcsv()
    .from.string(csv)
    .transform( function(row){
      row.unshift(row.pop());
      return row;
    })
    .on('record', function(row, index){

      if(index === 0) {
        header = row;
      }else{
        var obj = {};
        header.forEach(function(column, index) {
          obj[column.trim()] = row[index].trim();
        })
        record.push(obj);
      }
    })
    .on('end', function(count){

      if (transform && {}.toString.call(transform) === '[object Function]') {
        var buffer = transform(record);
        if (buffer) {
          record = buffer;
        } else {
          console.error('transform function error');
        }
      }
      // when writing to a file, use the 'close' event
      // the 'end' event may fire before the file has been written
      if(config.output !== null) {

        if (config.multiple && config.multiple === true){

          record.forEach(function(item,i){
              var split = config.output.split('.json')[0];
              var output = split+'-'+i+'.json';
              self.createFile(output,item,callback);
          });

        } else {
          self.createFile(config.output,record,callback);
        }
      }else {
        callback(null, record);
      }

    })
    .on('error', function(error){
      console.error(error.message);
    });
}

