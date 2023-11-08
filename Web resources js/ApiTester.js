// Online Javascript Editor for free
// Write, Edit and Run your Javascript code using JS Online Compiler

var timesheet = [
  {
    createdby: "Hao",
    new_duration: 8,
    new_date: "2023-10-17",
    new_currentstatus: "Office",
  },

  {
    createdby: "Dung",
    new_duration: 8,
    new_date: "2023-10-11",
    new_currentstatus: "Office",
  },

  {
    createdby: "Dung",
    new_duration: 3,
    new_date: "2023-10-11",
    new_currentstatus: "After22h",
  },
  {
    createdby: "Dung",
    new_duration: 8,
    new_date: "2023-10-12",
    new_currentstatus: "Office",
  },

  {
    createdby: "Dung",
    new_duration: 2,
    new_date: "2023-10-12",
    new_currentstatus: "After22h",
  },
  {
    createdby: "Hao",
    new_duration: 7,
    new_date: "2023-10-17",
    new_currentstatus: "Office",
  },
  {
    createdby: "Khang",
    new_duration: 8,
    new_date: "2023-08-11",
    new_currentstatus: "After22h",
  },
];

var vendoremployee = [
  {
    employee_name: "Hao",
    vendoridRef: "192134141df23",
  },
  {
    employee_name: "Dung",
    vendoridRef: "192134141df23",
  },
  {
    employee_name: "Khang",
    vendoridRef: "113df",
  },
];

var vendor = [
  {
    vendorid: "192134141df23",
    OTrate: 1,
    after22h: 3,
  },
  {
    vendorid: "113df",
    OTrate: 1,
    after22h: 3,
  },
];

function groupByDayAndCreatedBy(rawData, month, year) {
  var daysInMonth = new Date(year, month, 0).getDate();
  var totalOfficeHours_index = daysInMonth + 1
  var totalBefore22h_index = totalOfficeHours_index + 1
  var totalAfter22h_index = totalBefore22h_index + 1
  var result = [];

  // Tạo một object tạm để lưu trữ dữ liệu khi nhóm
  var temp = {};
  rawData.forEach(function (item) {
    // Lấy ngày từ trường 'new_date'
    var date = new Date(item.new_date);

    // Kiểm tra xem ngày này có thuộc tháng và năm đang xét hay không
    if (date.getMonth() + 1 === month && date.getFullYear() === year) {
      // Nếu chưa có dữ liệu cho 'createdby' này, khởi tạo một object mới
      if (!temp[item.createdby]) {
        temp[item.createdby] = {};
        for (var i = 1; i <= totalAfter22h_index; i++) {
          temp[item.createdby][i] = null;
        }
      }

      // Cộng dồn 'new_duration' vào ngày tương ứng trong kết quả
      if (temp[item.createdby][date.getDate()]) {
        temp[item.createdby][date.getDate()] += item.new_duration;
      } else {
        temp[item.createdby][date.getDate()] = item.new_duration;
      }


      var status = item.new_currentstatus;
      if(status == "Office"){
        if (temp[item.createdby][totalOfficeHours_index]) {
        temp[item.createdby][totalOfficeHours_index] += item.new_duration;
        } else {
        temp[item.createdby][totalOfficeHours_index] = item.new_duration;
        }
      }
      if(status == "After22h"){
        if (temp[item.createdby][totalAfter22h_index]) {
        temp[item.createdby][totalAfter22h_index] += item.new_duration;
        } else {
        temp[item.createdby][totalAfter22h_index] = item.new_duration;
        }
      }
      if(status == "Before22h"){
        if (temp[item.createdby][totalBefore22h_index]) {
        temp[item.createdby][totalBefore22h_index] += item.new_duration;
        } else {
        temp[item.createdby][totalBefore22h_index] = item.new_duration;
        }
      }
    }
  });

  console.log(temp)

  // Chuyển dữ liệu từ object tạm sang mảng kết quả
  for (var key in temp) {
    var obj = { createdby: key };
    for (var i = 1; i <= daysInMonth; i++) {
      obj[i] = temp[key][i];
    }
    obj["TotalOfficeHours"] = temp[key][totalOfficeHours_index];
    obj["TotalBefore22h"] = temp[key][totalBefore22h_index];
    obj["TotalAfter22h"] = temp[key][totalAfter22h_index];
    result.push(obj);
  }
  return result;
}

var result = timesheet.map((ts) => {
  var ve = vendoremployee.find((ve) => ve.employee_name === ts.createdby);
  var v = vendor.find((v) => v.vendorid === ve.vendoridRef);
  return { ...ts, ...ve, ...v };
});

// console.log(timesheet);

console.log(groupByDayAndCreatedBy(timesheet, 10, 2023));
