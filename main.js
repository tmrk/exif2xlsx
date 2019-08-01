'use strict';

var listData = [];
var wrapper = document.createElement('div'),
    wrapperText = document.createElement('div'),
    output = document.createElement('div');
    wrapper.id = 'wrapper';
    output.id = 'output';
var files = document.createElement('input');
    files.setAttribute('type', 'file');
    files.setAttribute('multiple', '');
    wrapper.appendChild(files);
    wrapper.appendChild(wrapperText);

function s(number) {
  if (number < 2) return '';
  else return 's';
}

wrapperText.innerHTML = '<p>Drag files here or click here to select</p>';

    files.addEventListener('change', function() {
      wrapperText.innerHTML = '<p>'+ files.files.length +' file' + s(files.files.length) + ' selected</p>';
      wrapper.classList.add('on');
      exportButton.classList.add('on');
      var filesList = this.files;
      for (var i = 0; i < filesList.length; i++) {
        (function () {
          var file = filesList[i];
          var reader = new FileReader();
          var fileData = { 'File name': file.name };
          reader.addEventListener('load', function(e) {
            var exif = EXIF.readFromBinaryFile(this.result),
                exifLat = exif.GPSLatitude,
                exifLong = exif.GPSLongitude, latitude, longitude;
            if (exif.GPSLatitudeRef == 'S') latitude = (exifLat[0]*-1) + (( (exifLat[1]*-60) + (exifLat[2]*-1) ) / 3600);
            else latitude = exifLat[0] + (( (exifLat[1]*60) + exifLat[2] ) / 3600);
            if (exif.GPSLongitudeRef == 'W') longitude = (exifLong[0]*-1) + (( (exifLong[1]*-60) + (exifLong[2]*-1) ) / 3600);
            else longitude = exifLong[0] + (( (exifLong[1]*60) + exifLong[2] ) / 3600);
            fileData['Date of photo'] = exif.DateTimeOriginal.slice(0, 10).replace(/:/g, '-');
            fileData['Time of photo'] = exif.DateTimeOriginal.slice(-8);
            fileData['Latitude'] = latitude;
            fileData['Longitude'] = longitude;
            fileData['Altitude'] = exif.GPSAltitude * 1;
            listData.push(fileData);
          });
          reader.readAsArrayBuffer(file);
        }());
      }
    });

    files.addEventListener('dragenter', function() {
      wrapper.classList.add('dragging');
    });
    files.addEventListener('drop', function() {
      wrapper.classList.remove('dragging');
    });
    files.addEventListener('dragleave', function() {
      wrapper.classList.remove('dragging');
    });

var exportButton = document.createElement('button');
    exportButton.innerHTML = 'Export to Excel';
    exportButton.addEventListener('click', function() {
      var ws = XLSX.utils.json_to_sheet(listData);
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      var wbout = XLSX.write(wb, {bookType:'xlsx', type:'binary'});
      XLSX.writeFile(wb, 'Export_' + new Date().toISOString().replace('T', '_').replace(/:/g, '').slice(0,-5) + '.xlsx');
    })

document.addEventListener('DOMContentLoaded', function(event) {
    document.body.appendChild(wrapper);
    document.body.appendChild(exportButton);
    document.body.appendChild(output);
});
