'use strict';

var listData = [];
var filedrag = document.createElement('div');
    filedrag.id = 'filedrag';
var files = document.createElement('input');
    files.setAttribute('type', 'file');
    files.setAttribute('multiple', '');
    files.addEventListener('change', function() {
      for (var i = 0; i < this.files.length; i++) {
        var file = this.files[i];
        var reader = new FileReader();
        reader.addEventListener('load', function(e) {
          var exif = EXIF.readFromBinaryFile(this.result),
              exifLat = exif.GPSLatitude,
              exifLong = exif.GPSLongitude, latitude, longitude;
          if (exif.GPSLatitudeRef == 'S') latitude = (exifLat[0]*-1) + (( (exifLat[1]*-60) + (exifLat[2]*-1) ) / 3600);
          else latitude = exifLat[0] + (( (exifLat[1]*60) + exifLat[2] ) / 3600);
          if (exif.GPSLongitudeRef == 'W') longitude = (exifLong[0]*-1) + (( (exifLong[1]*-60) + (exifLong[2]*-1) ) / 3600);
          else longitude = exifLong[0] + (( (exifLong[1]*60) + exifLong[2] ) / 3600);
          var fileData = {
            'File name':       file.name,
            'Date of photo':   exif.DateTimeOriginal.slice(0, 10).replace(/:/g, '-'),
            'Time of photo':   exif.DateTimeOriginal.slice(-8),
            'Latitude':        latitude,
            'Longitude':       longitude,
            'Altitude':        exif.GPSAltitude * 1
          };
          listData.push(fileData);
        })
        reader.readAsArrayBuffer(file);
      }
    });

var exportButton = document.createElement('button');
    exportButton.innerHTML = 'Export';
    exportButton.addEventListener('click', function() {
      var ws = XLSX.utils.json_to_sheet(listData);
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "sheetname");
      var wbout = XLSX.write(wb, {bookType:'xlsx', type:'binary'});
      XLSX.writeFile(wb, 'filename.xlsx');
    })

document.addEventListener('DOMContentLoaded', function(event) {
    document.body.appendChild(files);
    document.body.appendChild(exportButton);
});
