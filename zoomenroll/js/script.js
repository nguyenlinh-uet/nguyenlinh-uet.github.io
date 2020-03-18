const today = new Date();    
const year = today.getFullYear();
const month = today.getMonth();
const day = today.getDay();
const regex = /([01]\d|2[0123]):([012345]\d):([012345]\d)\t From  (.+?)_(.+?) :/;
const myRe = new RegExp(regex);
function getResult() {
    var data = $("#content")[0].value;
    var allline = data.split("\n");
    
    if(allline.length==1&&allline[0]==""){
        alert('Nội dung không được để trống');
        //return;
    };

    var start_time_enrollment;
    var end_time_enrollment;
    try{
        var start_time = $('#start_time')[0].value.split(':');
        var end_time = $('#end_time')[0].value.split(':');
        start_time_enrollment = new Date(year, month, day, start_time[0], start_time[1], start_time[2], 0); // giờ băý đầu điểm danh
        end_time_enrollment = new Date(year, month, day, end_time[0],end_time[1], end_time[2], 0); // giờ kết thúc điểm danh
    }catch(err){
        alert('Ngày giờ không hợp lệ');
        return;
    }
    var valid_enroll = []; // list điểm danh hợp lệ
    var valid_student_id = []; // dùng để check đã điểm danh student chưa

    allline.forEach(line => {
        if (!line.includes("(Privately)")) {
            // lọc những bình luận không phải riêng tư
            var founded = myRe.exec(line);
            if (founded) {
                let now = new Date(
                    year,
                    month,
                    day,
                    founded[1],
                    founded[2],
                    founded[3],
                    0
                ); // get time comment

                if (now >= start_time_enrollment && now <= end_time_enrollment) {
                    // giờ comment hợp lệ
                    let student = {
                        student_id: founded[4],
                        name: founded[5].trim()
                    };
                    if (!valid_student_id.includes(student.student_id)) {
                        // nếu chưa lưu vào danh sách điểm danh
                        valid_enroll.push(student);
                        valid_student_id.push(student.student_id);
                    }
                }
            }
        }
    });
    var wb = XLSX.utils.book_new();
    wb.Props = {
        Title: "Danh sách điểm danh - Zoom Meeting",
        Subject: "Danh sách điểm danh - Zoom Meeting",
        Author: "NguyenLinhUET",
        CreatedDate: new Date()
    };
    wb.SheetNames.push("danhsachdiemdanh");
    var ws_data = [['MSSV' , 'Họ và tên']];
    
    console.log(`Tổng số sinh viên đã điểm danh: ${valid_enroll.length}`);
    console.log("MSSV      | Họ và tên");
    valid_enroll.forEach(student => {
        console.log(`${student.student_id}  | ${student.name}`);
        ws_data.push([student.student_id,student.name]);
    });
    ws_data.push(['Đã điểm danh:',valid_enroll.length]);
    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    wb.Sheets["danhsachdiemdanh"] = ws;
    var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'danhsachdiemdanh.xlsx');
}

function s2ab(s) { 
    var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
    var view = new Uint8Array(buf);  //create uint8array as viewer
    for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
    return buf;    
}