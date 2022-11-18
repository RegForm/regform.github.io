let selectedFile;
let totalInfo = []; // file contents EXCEL
//let nStud = document.getElementById('nStud1')

// console.log(window.XLSX)


document.getElementById('excel').addEventListener("change",(event)=>{
    selectedFile = event.target.files[0];
})

function findInfo(id) {

    let personFinded = 0

    id = parseInt(id.match(/\d+/)) // if id="find1" => id=1
    let nStud = document.querySelector('#nStud'+id)

    let fStud = document.querySelector('#sFind'+id)

    let dateUntil = document.querySelector("#dateUntil"+id)
    let purpose = document.querySelector("#purpose"+id)
    let grazd = document.querySelector("#grazd"+id)
    let faculty = document.querySelector("#faculty"+id)
    let levelEducation = document.querySelector("#levelEducation"+id)
    let course = document.querySelector("#course"+id)
    let numOrder = document.querySelector("#numOrder"+id)
    let orderFrom = document.querySelector("#orderFrom"+id)
    let orderUntil = document.querySelector("#orderUntil"+id)
    let typeFunding = document.querySelector("#typeFunding"+id)
    let numContract = document.querySelector("#numContract"+id)
    let contractFrom = document.querySelector("#contractFrom"+id)
    let lastNameRu = document.querySelector("#lastNameRu"+id)
    let firstNameRu = document.querySelector("#firstNameRu"+id)
    let patronymicRu = document.querySelector("#patronymicRu"+id)
    let lastNameEn = document.querySelector("#lastNameEn"+id)
    let firstNameEn = document.querySelector("#firstNameEn"+id)
    let dateOfBirth = document.querySelector("#dateOfBirth"+id)
    let gender = document.querySelector("#gender"+id)
    let documentPerson = document.querySelector("#documentPerson"+id)
    let placeStateBirth = document.querySelector("#placeStateBirth"+id)
    let series = document.querySelector("#series"+id)
    let idPassport = document.querySelector("#idPassport"+id)
    let dateOfIssue = document.querySelector("#dateOfIssue"+id)
    let validUntil = document.querySelector("#validUntil"+id)
    let typeVisa = document.querySelector("#typeVisa"+id)
    let seriesVisa = document.querySelector("#seriesVisa"+id)
    let idVisa = document.querySelector("#idVisa"+id)
    let dateOfIssueVisa = document.querySelector("#dateOfIssueVisa"+id)
    let validUntilVisa = document.querySelector("#validUntilVisa"+id)
    let identifierVisa = document.querySelector("#identifierVisa"+id)
    let numInvVisa = document.querySelector("#numInvVisa"+id)
    let seriesMigration = document.querySelector("#seriesMigration"+id)
    let idMigration = document.querySelector("#idMigration"+id)
    let dateArrivalMigration = document.querySelector("#dateArrivalMigration"+id)
    let homeAddress = document.querySelector("#homeAddress"+id)
    let addressHostel = document.querySelector("#addressHostel"+id)
    let numRoom = document.querySelector("#numRoom"+id)
    let numRental = document.querySelector("#numRental"+id)
    let addressResidence = document.querySelector("#addressResidence"+id)
    let infHost = document.querySelector("#infHost"+id)
    let phone = document.querySelector("#phone"+id)
    let mail = document.querySelector("#mail"+id)
    let notificationFrom = document.querySelector("#notificationFrom"+id)
    let notificationUntil = document.querySelector("#notificationUntil"+id)
    let issuedBy = document.querySelector("#issuedBy"+id)






    let deleteButton = document.querySelector("#deleteButton"+id)


    // document.getElementById('find1').addEventListener('click',()=> {
    if (selectedFile && nStud.value!=0){
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event)=>{
            let data = event.target.result;
            let workbook = XLSX.read(data,{type:"binary"});

            let sheet = Object.keys(workbook.Sheets)[0]
            let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
            totalInfo = JSON.stringify(rowObject, undefined, 1)

            totalInfo = JSON.parse(totalInfo)

            // console.log(totalInfo)
            for (let i = 0; i<totalInfo.length; i++) {
                if (totalInfo[i]['Порядковый номер'] == nStud.value) {

                    personFinded = 1 // for print error

                    let dateStart = new Date(1900,0,0)
                    let dateEnd = new Date(dateStart)



                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['СРОК ОБУЧЕНИЯ ДО'])
                    dateUntil.value = dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    //purpose.value	=	totalInfo[i]['']
                    switch (totalInfo[i]['Гражданство (подданство)/ Citizenship']) {
                        case 'Китай/ China':
                            grazd.value = "Китай"
                            break
                        case 'Вьетнам/Vietnam':
                            grazd.value = "Вьетнам"
                            break
                        case 'Монголия/Mongolia':
                            grazd.value = "Монголия"
                            break
                        case 'Туркменистан /Turkmenistan':
                            grazd.value = "Туркменистан"
                            break
                        case 'Казахстан/Kazakhstan':
                            grazd.value ="Казахстан"
                            break
                        case 'Узбекистан/Uzbekistan':
                            grazd.value ="Узбекистан"
                            break
                        case 'Таджикистан/Tadjikistan':
                            grazd.value ="Таджикистан"
                            break
                        case 'Украина/Ukraine':
                            grazd.value ="Украина"
                            break
                        case 'Украина (ЛНР)/Ukraine(LNR)':
                            grazd.value ="Украина (ЛНР)"
                            break
                        case 'Украина (ДНР)/Ukraine(DNR)':
                            grazd.value ="Украина (ДНР)"
                            break
                        default:
                            grazd.value = totalInfo[i]['Гражданство (подданство)/ Citizenship']
                            break
                    }




                    faculty.value	=	totalInfo[i]['Институт & Факультет  / Institute & Faculty ']
                    levelEducation.value	=	totalInfo[i]['УРОВЕНЬ ОБРАЗОВАНИЯ/ LEVEL OF EDUCATION']
                    course.value	=	totalInfo[i]['КУРС ОБУЧЕНИЯ/YEAR OF STUDYING']
                    numOrder.value	=	totalInfo[i]['№ Приказа']

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Зачислен Приказом от'])
                    orderFrom.value = dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['СРОК ОБУЧЕНИЯ ДО'])
                    orderUntil.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)



                    typeFunding.value	=	totalInfo[i]['Тип финансирования/Type of Funding (state funded / paid tuition)']
                    numContract.value	=	totalInfo[i]['№ ДОГОВОРА ОБ ОКАЗАНИИ ПЛАТНЫХ ОБРАЗОВАТЕЛЬНЫХ УСЛУГ']

                    if (totalInfo[i]['Договор от']>20000) {
                        let dateStartCont = new Date(1900,0,-1)
                        let dateEndCont = new Date(dateStart)
                        dateEndCont.setDate(dateStartCont.getDate()+totalInfo[i]['Договор от'])
                        contractFrom.value = dateEndCont.toLocaleDateString()
                        dateEndCont = new Date(dateStartCont)

                    }
                    else {
                        contractFrom.value	=	totalInfo[i]['Договор от']
                    }


                    lastNameRu.value	=	totalInfo[i]['ФАМИЛИЯ (На русском языке) /SECOND NAME (in Russian)']
                    firstNameRu.value	=	totalInfo[i]['ИМЯ  (На русском языке) / FIRST NAME (in Russian)']
                    patronymicRu.value	=	totalInfo[i]['ОТЧЕСТВО  (На русском языке) ']
                    lastNameEn.value	=	totalInfo[i]['ФАМИЛИЯ (На английском языке)/ SECOND NAME (in English)']
                    firstNameEn.value	=	totalInfo[i]['ИМЯ  (На английском языке) / FIRST NAME (in English)']

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Год рождения / Date of birth'])
                    dateOfBirth.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    gender.value	=	totalInfo[i]['Пол / Sex']
                    // documentPerson.value	=	totalInfo[i]['ДОКУМЕНТ, УДОСТОВЕРЯЮЩИЙ ЛИЧНОСТЬ/IDENTITY DOCUMENT']
                    placeStateBirth.value	=	totalInfo[i]['Место рождения (Страна, город) / Place of birth (Country, city/town)']
                    series.value	=	totalInfo[i]['СЕРИЯ ПАСПОРТА/PASSPORT SERIES *']
                    idPassport.value	=	totalInfo[i]['НОМЕР ПАСПОРТА № /  PASSPORT NUMBER № *']

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Дата выдачи / Date of issue'])
                    dateOfIssue.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Срок действия (если есть) / Date of expiry'])
                    validUntil.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    typeVisa.value	=	totalInfo[i]['ВИД И РЕКВИЗИТЫ ДОКУМЕНТА, ПОДТВЕРЖДАЮЩЕГО ПРАВО НА ПРЕБЫВАНИЕ (ПРОЖИВАНИЕ) В РОССИЙСКОЙ ФЕДЕРАЦИИ ']
                    seriesVisa.value	=	totalInfo[i]['СЕРИЯ ВИЗЫ/VISA SERIES *']
                    idVisa.value	=	totalInfo[i]['НОМЕР ВИЗЫ №/ VISA NUMBER № *']

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Дата выдачи / Date of issue *'])
                    dateOfIssueVisa.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Срок действия / Date of expiry *'])
                    validUntilVisa.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    identifierVisa.value	=	totalInfo[i]['Идентификатор визы/ Visa ID №']
                    numInvVisa.value	=	totalInfo[i]['№ приглашения']
                    seriesMigration.value	=	totalInfo[i]['СЕРИЯ МИГРАЦИОННОЙ КАРТЫ/ MIGRATION CARD SERIES']
                    idMigration.value	=	totalInfo[i]['№ МИГРАЦИОННОЙ КАРТЫ/ MIGRATION CARD NUMBER']

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Срок пребывания: С /Duration of stay: From'])
                    dateArrivalMigration.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    homeAddress.value	=	totalInfo[i]["АДРЕС В СТРАНЕ ПОСТОЯННОГО ПРОЖИВАНИЯ (НА РОДИНЕ)\r\n1)Cтрана/Country of origin\r\n2)Провинция (или область) / Province\r\n3)Город / City \r\n4)Улица / Street\r\n5)№ дома / building №\r\n6)№ Квартиры / Apt №"]

                    addressHostel.value	=	totalInfo[i]['АДРЕС ПРОЖИВАНИЯ (ОБЩЕЖИТИЕ)']
                    numRoom.value	=	totalInfo[i]['№ КОМНАТЫ В ОБЩЕЖИТИИ МПГУ *']
                    numRental.value	=	totalInfo[i]['№ Договора найма *']
                    addressResidence.value	=	totalInfo[i]['АДРЕС ПРОЖИВАНИЯ В КВАРТИРЕ/ОТЕЛЕ:']
                    infHost.value	=	totalInfo[i]['СВЕДЕНИЯ О ПРИНИМАЮЩЕЙ СТОРОНЕ ( ЕСЛИ ВЫ ЖИВЕТЕ В КВАРТИРЕ)']
                    phone.value	=	totalInfo[i]['Номер телефона/Phone number ']
                    mail.value	=	totalInfo[i]['Ваш E-mail ']

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['УВЕДОМЛЕНИЕ О ПРИБЫТИИ ИНОСТРАННОГО ГРАЖДАНИНА С ...'])
                    notificationFrom.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)

                    dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Срок действия / Date of expiry *'])
                    notificationUntil.value	=	dateEnd.toISOString().split('T')[0]
                    dateEnd = new Date(dateStart)


                    issuedBy.value	=	totalInfo[i]['УВЕДОМЛЕНИЕ О ПРИБЫТИИ ИНОСТРАННОГО ГРАЖДАНИНА (КЕМ ВЫДАН ДОКУМЕНТ)']





                    // dateEnd.setDate(dateStart.getDate()+totalInfo[i]['Год рождения / Date of birth '])
                    // dateOfBirth.value = dateEnd.toLocaleDateString()
                    // dateEnd = new Date(dateStart)
                    //
                    // //series.value = totalInfo[i]['']
                    // idPassport.value = totalInfo[i]['№ паспорта / Passport №']







                    if (levelEducation.value == undefined|| levelEducation.value == '' || levelEducation.value == ' ') {
                        levelEducation.value = ' '
                        levelEducation.text = ''
                        levelEducation.selected = true
                    }
                    if (faculty.value == undefined || faculty.value == '' || faculty.value == ' ') {
                        faculty.value = ' '
                        faculty.text = ''
                        faculty.selected = true
                    }
                    if (course.value == undefined|| course.value == '' || course.value == ' ') {
                        course.value = ' '
                        course.text = ''
                        course.selected = true
                    }
                }
            }
            if (personFinded == 0) {
                alert('Студент с номером '+nStud.value+ ' не найден')
            }

        }
    }

    // });
}




function updateNameDisplay() {
    var input = document.querySelector('#excel');
    var preview = document.querySelector('.preview');
    var fileTypes = [
        'application/excel',
        'application/vnd.ms-excel',
        'application/x-excel',
        'application/x-msexcel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ]
    var curFiles = input.files;


    while(preview.firstChild) {
        preview.removeChild(preview.firstChild);
    }

    if(curFiles.length === 0) {
        var para = document.createElement('p');
        para.textContent = 'Вы не выбрали файл';
        preview.appendChild(para);
    } else {

        var para = document.createElement('p');
        if(validFileType(curFiles[0])) {
            para.textContent = 'File name ' + curFiles[0].name;
            var image = document.createElement('img');
            image.className = 'iconFile'
            image.src = '../excel.png';

            preview.appendChild(image);
            preview.appendChild(para);

        } else {
            para.textContent = 'Файл ' + curFiles[0].name + ' имеет неверный формат.';
            preview.appendChild(para);
        }

        // list.appendChild(preview);

    }



    function validFileType(file) {
        for(var i = 0; i < fileTypes.length; i++) {
            if(file.type === fileTypes[i]) {
                return true;
            }
        }

        return false;
    }

}





