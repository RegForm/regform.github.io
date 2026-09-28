let selectedFile;
let totalInfo = [];
let excelCachePromise = null;

document.getElementById('excel').addEventListener('change', event => {
    selectedFile = event.target.files[0];
    excelCachePromise = selectedFile ? readExcel(selectedFile, PREFERRED_SHEET, NUMBER_COLUMN) : null;
    // A rejected promise is handled when Find is clicked; prevent unhandled rejection meanwhile.
    if (excelCachePromise) excelCachePromise.catch(() => {});
});
const PREFERRED_SHEET = 'Ответы на форму (1)';
const NUMBER_COLUMN = 'Порядковый номер';

async function findInfo(id) {


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


    if (!excelCachePromise || !nStud.value.trim()) {
        alert('Выберите Excel и укажите номер студента');
        return;
    }
    try {
        const parsed = await excelCachePromise;
        totalInfo = parsed.rows;
        const row = parsed.byNumber.get(String(nStud.value).trim());
        if (!row) {
            alert('Студент с номером ' + nStud.value + ' не найден');
            return;
        }









                    //purpose.value	=	row['']
                    switch (row['Гражданство (подданство)/ Citizenship']) {
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
                            grazd.value = row['Гражданство (подданство)/ Citizenship']
                            break
                    }




                    if (excelDateToISO(row['Зачислен Приказом от'], parsed.date1904)) {

                        orderFrom.value = excelDateToISO(row['Зачислен Приказом от'], parsed.date1904)
                    }
                    else {orderFrom.value = ''}

                    if (excelDateToISO(row['СРОК ОБУЧЕНИЯ ДО'], parsed.date1904)) {

                        orderUntil.value = excelDateToISO(row['СРОК ОБУЧЕНИЯ ДО'], parsed.date1904)
                    }
                    else {
                        orderUntil.value = ''}

                    /* new */
                    if (excelDateToISO(row['ПО'], parsed.date1904)) {

                        dateUntil.value = excelDateToISO(row['ПО'], parsed.date1904)
                    }
                    else {
                        dateUntil.value = ''
                    }

                    /* !new */
                    if (excelDateToISO(row['Срок действия (если есть) / Date of expiry'], parsed.date1904)) {

                        validUntil.value = excelDateToISO(row['Срок действия (если есть) / Date of expiry'], parsed.date1904)
                    }
                    else {validUntil.value = ''}
                    if (excelDateToISO(row['Дата выдачи / Date of issue *'], parsed.date1904)) {

                        dateOfIssueVisa.value = excelDateToISO(row['Дата выдачи / Date of issue *'], parsed.date1904)
                    }
                    else {dateOfIssueVisa.value = ''}
                    if (excelDateToISO(row['Срок действия / Date of expiry *'], parsed.date1904)) {

                        validUntilVisa.value = excelDateToISO(row['Срок действия / Date of expiry *'], parsed.date1904)

                        notificationUntil.value = excelDateToISO(row['Срок действия / Date of expiry *'], parsed.date1904)
                    }
                    else {validUntilVisa.value = ''
                        notificationUntil.value = ''}
                    if (excelDateToISO(row['УВЕДОМЛЕНИЕ О ПРИБЫТИИ ИНОСТРАННОГО ГРАЖДАНИНА С ...'], parsed.date1904)) {

                        notificationFrom.value = excelDateToISO(row['УВЕДОМЛЕНИЕ О ПРИБЫТИИ ИНОСТРАННОГО ГРАЖДАНИНА С ...'], parsed.date1904)
                    }
                    else {notificationFrom.value = ''}

                    if (excelDateToISO(row['Срок пребывания: С /Duration of stay: From'], parsed.date1904)) {

                        dateArrivalMigration.value = excelDateToISO(row['Срок пребывания: С /Duration of stay: From'], parsed.date1904)
                    } else {dateArrivalMigration.value = ''}


                    if (excelDateToISO(row['Дата выдачи / Date of issue'], parsed.date1904)) {

                        dateOfIssue.value = excelDateToISO(row['Дата выдачи / Date of issue'], parsed.date1904)
                    }

                    else {dateOfIssue.value = ''}

                    if (excelDateToISO(row['Год рождения / Date of birth'], parsed.date1904)) {

                        dateOfBirth.value = excelDateToISO(row['Год рождения / Date of birth'], parsed.date1904)
                    }
                    else {dateOfBirth.value = ''}



                    faculty.value	=	row['Институт & Факультет  / Institute & Faculty ']
                    levelEducation.value	=	row['УРОВЕНЬ ОБРАЗОВАНИЯ/ LEVEL OF EDUCATION']
                    course.value	=	row['КУРС ОБУЧЕНИЯ/YEAR OF STUDYING']


                    numOrder.value	=	row['№ Приказа'] ? row['№ Приказа'] : ''







                    typeFunding.value	=	row['Тип финансирования/Type of Funding (state funded / paid tuition)'] ? row['Тип финансирования/Type of Funding (state funded / paid tuition)'] : ''
                    numContract.value	=	row['№ ДОГОВОРА ОБ ОКАЗАНИИ ПЛАТНЫХ ОБРАЗОВАТЕЛЬНЫХ УСЛУГ'] ? row['№ ДОГОВОРА ОБ ОКАЗАНИИ ПЛАТНЫХ ОБРАЗОВАТЕЛЬНЫХ УСЛУГ'] : ""

                    if (excelDateToISO(row['Договор от'], parsed.date1904)) {
                        contractFrom.value = formatDateRU(excelDateToISO(row['Договор от'], parsed.date1904))
                    }
                    else {
                        contractFrom.value	=	""
                    }


                    lastNameRu.value	=	row['ФАМИЛИЯ (На русском языке) /SECOND NAME (in Russian)'] ? row['ФАМИЛИЯ (На русском языке) /SECOND NAME (in Russian)'] : ""
                    firstNameRu.value	=	row['ИМЯ  (На русском языке) / FIRST NAME (in Russian)'] ? row['ИМЯ  (На русском языке) / FIRST NAME (in Russian)'] : ""
                    patronymicRu.value	=	row['ОТЧЕСТВО  (На русском языке) '] ? row['ОТЧЕСТВО  (На русском языке) '] : ""
                    lastNameEn.value	=	row['ФАМИЛИЯ (На английском языке)/ SECOND NAME (in English)'] ? row['ФАМИЛИЯ (На английском языке)/ SECOND NAME (in English)'] : ''
                    firstNameEn.value	=	row['ИМЯ  (На английском языке) / FIRST NAME (in English)'] ? row['ИМЯ  (На английском языке) / FIRST NAME (in English)'] : ''



                    gender.value	=	row['Пол / Sex']
                    // documentPerson.value	=	row['ДОКУМЕНТ, УДОСТОВЕРЯЮЩИЙ ЛИЧНОСТЬ/IDENTITY DOCUMENT']
                    placeStateBirth.value	=	row['Место рождения (Страна, город) / Place of birth (Country, city/town)'] ? row['Место рождения (Страна, город) / Place of birth (Country, city/town)'] : ""
                    series.value	=	row['СЕРИЯ ПАСПОРТА/PASSPORT SERIES *'] ? row['СЕРИЯ ПАСПОРТА/PASSPORT SERIES *'] : ""
                    idPassport.value	=	row['НОМЕР ПАСПОРТА № /  PASSPORT NUMBER № *'] ? row['НОМЕР ПАСПОРТА № /  PASSPORT NUMBER № *'] : ""






                    typeVisa.value	=	row['ВИД И РЕКВИЗИТЫ ДОКУМЕНТА, ПОДТВЕРЖДАЮЩЕГО ПРАВО НА ПРЕБЫВАНИЕ (ПРОЖИВАНИЕ) В РОССИЙСКОЙ ФЕДЕРАЦИИ ']
                    seriesVisa.value	=	row['СЕРИЯ ВИЗЫ/VISA SERIES *'] ? row['СЕРИЯ ВИЗЫ/VISA SERIES *'] : ''
                    idVisa.value	=	row['НОМЕР ВИЗЫ №/ VISA NUMBER № *'] ? row['НОМЕР ВИЗЫ №/ VISA NUMBER № *'] : ''





                    identifierVisa.value	=	row['Идентификатор визы/ Visa ID №'] ? row['Идентификатор визы/ Visa ID №'] : ''
                    numInvVisa.value	=	row['№ приглашения'] ? row['№ приглашения'] : ""
                    seriesMigration.value	=	row['СЕРИЯ МИГРАЦИОННОЙ КАРТЫ/ MIGRATION CARD SERIES'] ? row['СЕРИЯ МИГРАЦИОННОЙ КАРТЫ/ MIGRATION CARD SERIES'] : ""
                    idMigration.value	=	row['№ МИГРАЦИОННОЙ КАРТЫ/ MIGRATION CARD NUMBER'] ? row['№ МИГРАЦИОННОЙ КАРТЫ/ MIGRATION CARD NUMBER'] : ""



                    homeAddress.value	=	row["АДРЕС В СТРАНЕ ПОСТОЯННОГО ПРОЖИВАНИЯ (НА РОДИНЕ)\n1)Cтрана/Country of origin\n2)Провинция (или область) / Province\n3)Город / City \n4)Улица / Street\n5)№ дома / building №\n6)№ Квартиры / Apt №"] ?
                        row["АДРЕС В СТРАНЕ ПОСТОЯННОГО ПРОЖИВАНИЯ (НА РОДИНЕ)\n1)Cтрана/Country of origin\n2)Провинция (или область) / Province\n3)Город / City \n4)Улица / Street\n5)№ дома / building №\n6)№ Квартиры / Apt №"] : ""
                    
                    addressHostel.value	=	row['АДРЕС ПРОЖИВАНИЯ (ОБЩЕЖИТИЕ)'] ? row['АДРЕС ПРОЖИВАНИЯ (ОБЩЕЖИТИЕ)'] : ""
                    numRoom.value	=	row['№ КОМНАТЫ В ОБЩЕЖИТИИ МПГУ *'] ? row['№ КОМНАТЫ В ОБЩЕЖИТИИ МПГУ *'] : ""
                    numRental.value	=	row['№ Договора найма *'] ? row['№ Договора найма *'] : ""
                    addressResidence.value	=	row['АДРЕС ПРОЖИВАНИЯ В КВАРТИРЕ/ОТЕЛЕ:'] ? row['АДРЕС ПРОЖИВАНИЯ В КВАРТИРЕ/ОТЕЛЕ:'] :  ""
                    infHost.value	=	row['СВЕДЕНИЯ О ПРИНИМАЮЩЕЙ СТОРОНЕ ( ЗАПОЛНИТЕ ЭТО ПОЛЕ ТОЛЬКО ЕСЛИ ВЫ ЖИВЕТЕ В КВАРТИРЕ)'] ? row['СВЕДЕНИЯ О ПРИНИМАЮЩЕЙ СТОРОНЕ ( ЗАПОЛНИТЕ ЭТО ПОЛЕ ТОЛЬКО ЕСЛИ ВЫ ЖИВЕТЕ В КВАРТИРЕ)'] : ""
                    phone.value	=	row['Номер телефона/Phone number '] ? row['Номер телефона/Phone number '] : ""
                    mail.value	=	row['Ваш E-mail '] ? row['Ваш E-mail '] : ""






                    issuedBy.value	=	row['УВЕДОМЛЕНИЕ О ПРИБЫТИИ ИНОСТРАННОГО ГРАЖДАНИНА (КЕМ ВЫДАН ДОКУМЕНТ)'] ? row['УВЕДОМЛЕНИЕ О ПРИБЫТИИ ИНОСТРАННОГО ГРАЖДАНИНА (КЕМ ВЫДАН ДОКУМЕНТ)'] : ""
                    //
                    // //series.value = row['']
                    // idPassport.value = row['№ паспорта / Passport №']







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
                
    } catch (error) {
        console.error('Ошибка импорта Excel', error);
        alert('Не удалось прочитать Excel: ' + error.message);
    }
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
            image.src = 'excel.png';

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







