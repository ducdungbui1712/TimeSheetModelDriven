var timesheet = [
  {
    _createdby_value: "Hao",
    new_duration: 8,
    new_date: "2023-10-17",
    new_currentstatus: "Office",
  },
   {
    _createdby_value: "Hao",
    new_duration: 2,
    new_date: "2023-10-17",
    new_currentstatus: "After22h",
  },
  {
    _createdby_value: "Dung",
    new_duration: 8,
    new_date: "2023-10-11",
    new_currentstatus: "Before22h",
  },
  {
    _createdby_value: "Hao",
    new_duration: 7,
    new_date: "2023-10-17",
    new_currentstatus: "Office",
  },
  {
    _createdby_value: "Hao",
    new_duration: 7,
    new_date: "2023-10-17",
    new_currentstatus: "Before22h",
  },
  {
    _createdby_value: "Khang",
    new_duration: 8,
    new_date: "2023-08-11",
    new_currentstatus: "Before22h",
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
    newweekdayoffice: 1,
    newweekdaybefore22h: 1.5,
    newweekdayafter22h: 1.5,
    newweekendbefore22h:2,
    newweekendafter22h:2,
 
  },
  {
    vendorid: "113df",
    newweekdayoffice: 1,
    newweekdaybefore22h: 1.5,
    newweekdayafter22h: 1.5,
    newweekendbefore22h:2,
    newweekendafter22h:2,
  
  },
];



var result = timesheet.map((ts) => {
  var ve = vendoremployee.find((ve) => ve.employee_name === ts._createdby_value);
  var v = vendor.find((v) => v.vendorid === ve.vendoridRef);
  return { ...ts, ...ve, ...v };
});

console.log(result)

function OtHour(item){
	let value = 1;
	let date = new Date(item.new_date);
	let day = date.getDay();
	if(day >= 1 && day <= 5 && item.new_currentstatus == "Office")
		value = item.new_duration * item.newweekdayoffice;
	else if (day >= 1 && day <= 5 && item.new_currentstatus == "Before22h")
		value = item.new_duration * item.newweekdaybefore22h;
	else if (day >= 1 && day <= 5 && item.new_currentstatus == "After22h")
		value = item.new_duration * item.newweekdayafter22h;
	else if ((day === 0 || day === 6) && item.new_currentstatus == "Before22h")
		value = item.new_duration * item.newweekendbefore22h;
	else if ((day === 0 || day === 6) && item.new_currentstatus == "After22h")
		value = item.new_duration * item.newweekendafter22h;
	return value;
}



function groupByDayAnd_createdby_value(rawData, month, year) {
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
      // Nếu chưa có dữ liệu cho '_createdby_value' này, khởi tạo một object mới
      if (!temp[item._createdby_value]) {
        temp[item._createdby_value] = {};
        for (var i = 1; i <= totalAfter22h_index; i++) {
          temp[item._createdby_value][i] = null;
        }
      }

      // Cộng dồn 'new_duration' vào ngày tương ứng trong kết quả
      if (temp[item._createdby_value][date.getDate()]) {
        temp[item._createdby_value][date.getDate()] += item.new_duration;
      } else {
        temp[item._createdby_value][date.getDate()] = item.new_duration;
      }


      var status = item.new_currentstatus;
      if(status == "Office"){
        if (temp[item._createdby_value][totalOfficeHours_index]) {
        temp[item._createdby_value][totalOfficeHours_index] += OtHour(item);
        } else {
        temp[item._createdby_value][totalOfficeHours_index] = OtHour(item);
        }
      }
      if(status == "After22h"){
        if (temp[item._createdby_value][totalAfter22h_index]) {
        temp[item._createdby_value][totalAfter22h_index] += OtHour(item);
        } else {
        temp[item._createdby_value][totalAfter22h_index] = OtHour(item);
        }
      }
      if(status == "Before22h"){
        if (temp[item._createdby_value][totalBefore22h_index]) {
        temp[item._createdby_value][totalBefore22h_index] += OtHour(item);
        } else {
        temp[item._createdby_value][totalBefore22h_index] = OtHour(item);
        }
      }
    }
  });

  //console.log(temp)

  // Chuyển dữ liệu từ object tạm sang mảng kết quả
  for (var key in temp) {
    var obj = { Employee: key };
    for (var i = 1; i <= daysInMonth; i++) {
      obj[i] = temp[key][i];
    }
    obj["TotalOfficeHours"] = temp[key][totalOfficeHours_index];
    obj["TotalBefore22h"] = temp[key][totalBefore22h_index];
    obj["TotalAfter22h"] = temp[key][totalAfter22h_index];
    obj["TotalWorkingHours"] = obj["TotalOfficeHours"] + obj["TotalBefore22h"] + obj["TotalAfter22h"]
    result.push(obj);
  }
  return result;
}

console.log(groupByDayAnd_createdby_value(result,10,2023))



function groupByDayAnd_createdby_value(rawData, month, year) {
  var daysInMonth = new Date(year, month, 0).getDate();
  var result = [];

  // Tạo một object tạm để lưu trữ dữ liệu khi nhóm
  var temp = {};
  rawData.forEach(function(item) {
      // Lấy ngày từ trường 'new_date'
      var date = new Date(item.new_date);
      
      // Kiểm tra xem ngày này có thuộc tháng và năm đang xét hay không
      if (date.getMonth() + 1 === month && date.getFullYear() === year) {
          // Nếu chưa có dữ liệu cho '_createdby_value' này, khởi tạo một object mới
          if (!temp[item._createdby_value]) {
              temp[item._createdby_value] = { _createdby_value: item._createdby_value };
              for (var i = 1; i <= daysInMonth; i++) {
                  temp[item._createdby_value][i] = null;
              }
          }

          // Cộng dồn 'new_duration' vào ngày tương ứng trong kết quả
          if (temp[item._createdby_value][date.getDate()]) {
              temp[item._createdby_value][date.getDate()] += item.new_duration;
          } else {
              temp[item._createdby_value][date.getDate()] = item.new_duration;
          }
      }
  });

  // Chuyển dữ liệu từ object tạm sang mảng kết quả
  for (var key in temp) {
      result.push(temp[key]);
  }










