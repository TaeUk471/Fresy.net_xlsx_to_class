const xlsx = require('xlsx');

const filePath = fresynet_client.xlsx;
const workbook = xlsx.readFile(filePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

console.log(xlsx);

let posts = [];
let post = {};

for (let cell in worksheet) {
    const cellAsString = cell.toString();

    if (cellAsString[1] !== 'Stop') {
        switch (cellAsString[0]){
            case 'A' : 
            post.supplier = worksheet[cell].v;
            break;
            //공급자
            case 'B' : 
            post.businessNumber = worksheet[cell].v;
            break;
            //공급자사업자번호
            case 'C' : 
            post.emailAdderss  = worksheet[cell].v;
            break;
            //공급자 이메일
            case 'D' : 
            post.headOffice = worksheet[cell].v;
            break;
            //본사
            case 'E' : 
            post.collectionEngineer = worksheet[cell].v;
            break;
            //수거기사
            case 'F' : 
            post.dischargeCompany = worksheet[cell].v;
            break;
            //배출업체
            case 'G' : 
            post.duedate = worksheet[cell].v;
            break;
            //마감일
            case 'H' : 
            post.kind = worksheet[cell].v;
            break;
            //종류
            case 'I' : 
            post.unitPrice = worksheet[cell].v;
            break;
            //단가
            case 'J' : 
            post.registrationNumber = worksheet[cell].v;
            break;
            //사업자번호
            case 'K' : 
            post.billAddress = worksheet[cell].v;
            break;
            //계산서 주소
            case 'L' : 
            post.workplaceAddress = worksheet[cell].v;
            break;
            //사업장 주소
            case 'M' : 
            post.supervisor = worksheet[cell].v;
            break;
            //사업자
            case 'N' : 
            post.blank = worksheet[cell].v;
            break;
            //사업자
            case 'O' : 
            post.supervisorEmail = worksheet[cell].v;
            break;
            //사업자
            case 'P' : 
            post.phoneNumber = worksheet[cell].v;
            break;
            //사업자

        }
        posts.push(post);
            post = {};
    }
}

console.log(posts);
