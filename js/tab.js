// <!-- DELETE TAB -->
function removeDummy(el) {
    // index elem
    let elem = parseInt(el.parentNode.id.match(/\d+/));


    let tabElem = document.getElementById('tab' + elem)
    let liElem = document.getElementById('li-tab' + elem)

    tabElem.parentNode.removeChild(tabElem);
    liElem.parentNode.removeChild(liElem);

    // var elem = document.getElementById('home');
    // elem.parentNode.removeChild(elem);
    // var elem = document.getElementById('l1');
    // elem.parentNode.removeChild(elem);
}


// RENAME TAB
function renameTab(ident) {
    let newId = parseInt(ident.id.match(/\d+/))
    let iTabs = document.getElementById('li-tab'+newId)
    if (ident.value == "" || ident.value == " ") {
        iTabs.querySelector('a').textContent = 'Вкладка '+newId
    }
    else {
        iTabs.querySelector('a').textContent = ident.value
    }

}

<!-- CREATE TAB -->
function createTab() {

    let countTabs = countTab()
    let indexNewTab = lastTab()


    // copy li-tab1
    let liTab1 = document.querySelector("li#li-tab1");
    let newLiTab1 = liTab1.cloneNode(true);
    newLiTab1.className = ''
    newLiTab1.id = 'li-tab' + indexNewTab
    newLiTab1.querySelector('a').href = "#tab" + indexNewTab
    newLiTab1.querySelector('a').textContent = 'Вкладка ' + indexNewTab
    let navTab = document.querySelector('ul.nav-tabs');
    // past li-tab1
    navTab.insertBefore(newLiTab1,navTab.children[countTabs])



    // copy tab1
    let tabDiv = document.querySelector("div#tab1");
    let newTabDiv = tabDiv.cloneNode(true);
    newTabDiv.className='tab-pane fade'
    newTabDiv.id = 'tab' + indexNewTab
    let tapContentDiv = document.querySelector('div.tab-content')

    // change name of all input
    let nStud=newTabDiv.querySelector('#nStud1')
    nStud.id =  'nStud' + indexNewTab
    let sFind=newTabDiv.querySelector('#sFind1')
    sFind.id =  'sFind' + indexNewTab
    let dateUntil=newTabDiv.querySelector('#dateUntil1')
    dateUntil.id =  'dateUntil' + indexNewTab
    let purpose=newTabDiv.querySelector('#purpose1')
    purpose.id =  'purpose' + indexNewTab
    let grazd=newTabDiv.querySelector('#grazd1')
    grazd.id =  'grazd' + indexNewTab
    let faculty=newTabDiv.querySelector('#faculty1')
    faculty.id =  'faculty' + indexNewTab
    let levelEducation=newTabDiv.querySelector('#levelEducation1')
    levelEducation.id =  'levelEducation' + indexNewTab
    let course=newTabDiv.querySelector('#course1')
    course.id =  'course' + indexNewTab
    let numOrder=newTabDiv.querySelector('#numOrder1')
    numOrder.id =  'numOrder' + indexNewTab
    let orderFrom=newTabDiv.querySelector('#orderFrom1')
    orderFrom.id =  'orderFrom' + indexNewTab
    let orderUntil=newTabDiv.querySelector('#orderUntil1')
    orderUntil.id =  'orderUntil' + indexNewTab
    let typeFunding=newTabDiv.querySelector('#typeFunding1')
    typeFunding.id =  'typeFunding' + indexNewTab
    let numContract=newTabDiv.querySelector('#numContract1')
    numContract.id =  'numContract' + indexNewTab
    let contractFrom=newTabDiv.querySelector('#contractFrom1')
    contractFrom.id =  'contractFrom' + indexNewTab
    let lastNameRu=newTabDiv.querySelector('#lastNameRu1')
    lastNameRu.id =  'lastNameRu' + indexNewTab
    let firstNameRu=newTabDiv.querySelector('#firstNameRu1')
    firstNameRu.id =  'firstNameRu' + indexNewTab
    let patronymicRu=newTabDiv.querySelector('#patronymicRu1')
    patronymicRu.id =  'patronymicRu' + indexNewTab
    let lastNameEn=newTabDiv.querySelector('#lastNameEn1')
    lastNameEn.id =  'lastNameEn' + indexNewTab
    let firstNameEn=newTabDiv.querySelector('#firstNameEn1')
    firstNameEn.id =  'firstNameEn' + indexNewTab
    let dateOfBirth=newTabDiv.querySelector('#dateOfBirth1')
    dateOfBirth.id =  'dateOfBirth' + indexNewTab
    let gender=newTabDiv.querySelector('#gender1')
    gender.id =  'gender' + indexNewTab
    // let documentPerson=newTabDiv.querySelector('#documentPerson1')
    // documentPerson.id =  'documentPerson' + indexNewTab
    let placeStateBirth=newTabDiv.querySelector('#placeStateBirth1')
    placeStateBirth.id =  'placeStateBirth' + indexNewTab
    let series=newTabDiv.querySelector('#series1')
    series.id =  'series' + indexNewTab
    let idPassport=newTabDiv.querySelector('#idPassport1')
    idPassport.id =  'idPassport' + indexNewTab
    let dateOfIssue=newTabDiv.querySelector('#dateOfIssue1')
    dateOfIssue.id =  'dateOfIssue' + indexNewTab
    let validUntil=newTabDiv.querySelector('#validUntil1')
    validUntil.id =  'validUntil' + indexNewTab
    let typeVisa=newTabDiv.querySelector('#typeVisa1')
    typeVisa.id =  'typeVisa' + indexNewTab
    let seriesVisa=newTabDiv.querySelector('#seriesVisa1')
    seriesVisa.id =  'seriesVisa' + indexNewTab
    let idVisa=newTabDiv.querySelector('#idVisa1')
    idVisa.id =  'idVisa' + indexNewTab
    let dateOfIssueVisa=newTabDiv.querySelector('#dateOfIssueVisa1')
    dateOfIssueVisa.id =  'dateOfIssueVisa' + indexNewTab
    let validUntilVisa=newTabDiv.querySelector('#validUntilVisa1')
    validUntilVisa.id =  'validUntilVisa' + indexNewTab
    let identifierVisa=newTabDiv.querySelector('#identifierVisa1')
    identifierVisa.id =  'identifierVisa' + indexNewTab
    let numInvVisa=newTabDiv.querySelector('#numInvVisa1')
    numInvVisa.id =  'numInvVisa' + indexNewTab
    let seriesMigration=newTabDiv.querySelector('#seriesMigration1')
    seriesMigration.id =  'seriesMigration' + indexNewTab
    let idMigration=newTabDiv.querySelector('#idMigration1')
    idMigration.id =  'idMigration' + indexNewTab
    let dateArrivalMigration=newTabDiv.querySelector('#dateArrivalMigration1')
    dateArrivalMigration.id =  'dateArrivalMigration' + indexNewTab
    let homeAddress=newTabDiv.querySelector('#homeAddress1')
    homeAddress.id =  'homeAddress' + indexNewTab
    let addressHostel=newTabDiv.querySelector('#addressHostel1')
    addressHostel.id =  'addressHostel' + indexNewTab
    let numRoom=newTabDiv.querySelector('#numRoom1')
    numRoom.id =  'numRoom' + indexNewTab
    let numRental=newTabDiv.querySelector('#numRental1')
    numRental.id =  'numRental' + indexNewTab
    let addressResidence=newTabDiv.querySelector('#addressResidence1')
    addressResidence.id =  'addressResidence' + indexNewTab
    let infHost=newTabDiv.querySelector('#infHost1')
    infHost.id =  'infHost' + indexNewTab
    let phone=newTabDiv.querySelector('#phone1')
    phone.id =  'phone' + indexNewTab
    let mail=newTabDiv.querySelector('#mail1')
    mail.id =  'mail' + indexNewTab
    let notificationFrom=newTabDiv.querySelector('#notificationFrom1')
    notificationFrom.id =  'notificationFrom' + indexNewTab
    let notificationUntil=newTabDiv.querySelector('#notificationUntil1')
    notificationUntil.id =  'notificationUntil' + indexNewTab
    let issuedBy=newTabDiv.querySelector('#issuedBy1')
    issuedBy.id =  'issuedBy' + indexNewTab
    let deleteButton=newTabDiv.querySelector('#deleteButton1')
    deleteButton.id =  'deleteButton' + indexNewTab






    let aInput = [nStud,
        dateUntil,

        grazd,

        numOrder,
        orderFrom,
        orderUntil,

        numContract,
        contractFrom,
        lastNameRu,
        firstNameRu,
        patronymicRu,
        lastNameEn,
        firstNameEn,
        dateOfBirth,

        placeStateBirth,
        series,
        idPassport,
        dateOfIssue,
        validUntil,

        seriesVisa,
        idVisa,
        dateOfIssueVisa,
        validUntilVisa,
        identifierVisa,
        numInvVisa,
        seriesMigration,
        idMigration,
        dateArrivalMigration,
        homeAddress,

        numRoom,
        numRental,
        addressResidence,
        infHost,
        phone,
        mail,
        notificationFrom,
        notificationUntil,
        issuedBy,
    ]

    for (let i of aInput) {
        i.value = ""
    }

    //past tab1
    tapContentDiv.append(newTabDiv)

    //filling Select
    // fillingFaculty(selFaculty, faculty.id)



    // fillingSelect(selCourse, course.id)
    // fillingSelect(selPlaceStateGetVisa, placeStateGetVisa.id)
    // fillingSelect(selPlaceCityGetVisaOther, placeCityGetVisa.id)
    // fillingSelect(selStateBirth,stateBirth.id)
    //
    //
    // fillingTypeVisa() //!!!
    //
    // fillingLevelEducation(selLevelEducation, levelEducation.id)
    // fillingGender(selGender, gender.id)

}







// last exist index+1
function lastTab() {
    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
    let lastElem = tabs[tabs.length-2]
    // let lastIndexTab = Number(lastElem.id[lastElem.id.length-1])
    let lastIndexTab = parseInt(lastElem.id.match(/\d+/))
    return lastIndexTab+1
}

// count tabs
function countTab() {
    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
    let lastElem = tabs.length-1
    return lastElem
}







// SELECT
// let selGrazd = ["","Китай", "Вьетнам", "Туркменистан", "Монголия", "Тайвань (Китай)", "Другое"]




// 100%
// fillingSelect(selCourse,'course1')
// fillingSelect(selPlaceStateGetVisa, 'placeStateGetVisa1')
// fillingSelect(selPlaceCityGetVisaOther, 'placeCityGetVisa1')
// fillingSelect(selStateBirth,'stateBirth1')
// fillingTypeVisa()

// fillingGender(selGender, 'gender1')
// fillingLevelEducation(selLevelEducation, 'levelEducation1')
// fillingFaculty(selFaculty, 'faculty1')

// function fillingTypeVisa() {
//     let chPurpose = document.getElementById('purpose')
//
//     let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
//
//     if (chPurpose.value == 'Учеба' || chPurpose.value == "Краткосрочная учеба") {
//         for (let i=0; i<countTab(); i++) {
//             let indexTab = parseInt(tabs[i].id.match(/\d+/))
//             fillingSelect(selTypeVisaUch, 'typeVisa'+indexTab)
//         }
//     }
//     else if (chPurpose.value == "(НТС)") {
//         for (let i=0; i<countTab(); i++) {
//             let indexTab = parseInt(tabs[i].id.match(/\d+/))
//             fillingSelect(selTypeVisaGum, 'typeVisa'+indexTab)
//         }
//     }
//     else if (chPurpose.value == "Трудовая деятельность") {
//         for (let i=0; i<countTab(); i++) {
//             let indexTab = parseInt(tabs[i].id.match(/\d+/))
//             fillingSelect(selTypeVisaRab, 'typeVisa'+indexTab)
//         }
//     }
// }

// function fillingSelect (nameMassiveSelect, idSelect) {
//     document.getElementById(idSelect).innerHTML=''
//
//     for (let i=0; i<nameMassiveSelect.length; i++) {
//         let opt = nameMassiveSelect[i]
//         let newOption = document.createElement('option')
//         let chooseSelect = document.getElementById(idSelect)
//         newOption.text = opt
//         newOption.value = opt
//         if (opt=="") {
//             newOption.selected = true
//             newOption.hidden = true
//             newOption.disabled = true
//         }
//         chooseSelect.add(newOption)
//     }
// }

// function fillingGender(nameMassiveSelect, idSelect) {
//     document.getElementById(idSelect).innerHTML=''
//     for (let i=0; i<nameMassiveSelect.length; i++) {
//         let opt = nameMassiveSelect[i]
//         let newOption = document.createElement('option')
//         let chooseSelect = document.getElementById(idSelect)
//         if (opt == "Мужской") {
//             newOption.value = "Мужской / Male"
//             newOption.text = opt
//         } else if (opt == "Женский") {
//             newOption.value = 'Женский / Female'
//             newOption.text = opt
//         }
//         else if (opt=="") {
//             newOption.selected = true
//             newOption.hidden = true
//             newOption.disabled = true
//         }
//         chooseSelect.add(newOption)
//     }
// }

// function fillingLevelEducation(nameMassiveSelect, idSelect) {
//     document.getElementById(idSelect).innerHTML=''
//     for (let i=0; i<nameMassiveSelect.length; i++) {
//         let opt = nameMassiveSelect[i]
//         let newOption = document.createElement('option')
//         let chooseSelect = document.getElementById(idSelect)
//         if (opt == "Подфак") {
//             newOption.value = "Подготовительный факультет (изучаю русский язык)/ The preparatory faculty"
//             newOption.text = opt
//         } else if (opt == "Бакалавриат") {
//             newOption.value = 'бакалавриат/bachelor degree'
//             newOption.text = opt
//         }
//         else if (opt == "Магистратура") {
//             newOption.value = 'магистратура/master degree'
//             newOption.text = opt
//         }
//         else if (opt == "Аспирантура") {
//             newOption.value = 'аспирантура/post-graduate studies'
//             newOption.text = opt
//         }
//         else if (opt == "Стажировка") {
//             newOption.value = 'Стажировка'
//             newOption.text = opt
//         }
//         else if (opt == "Стажировка(межвуз)") {
//             newOption.value = 'Стажировка(межвуз)'
//             newOption.text = opt
//         }
//
//         else if (opt=="") {
//             newOption.selected = true
//             newOption.hidden = true
//             newOption.disabled = true
//         }
//         chooseSelect.add(newOption)
//     }
// }


// on change select


// placeStateGetVisa
// function chPlaceStateGetVisa(ident) {
//     let indexTab = parseInt(ident.id.match(/\d+/))
//     if (ident.value == 'Китай') {
//         fillingSelect(selPlaceCityGetVisaChina, 'placeCityGetVisa'+ indexTab)
//     }
//     else if (ident.value == 'Турция') {
//         fillingSelect(selPlaceCityGetVisaTurkey, 'placeCityGetVisa'+ indexTab)
//     }
//     else if (ident.value == 'Туркменистан') {
//         fillingSelect(selPlaceCityGetVisaTurkmen, 'placeCityGetVisa'+ indexTab)
//     }
//     else if (ident.value == 'Другое') {
//         fillingSelect(selPlaceCityGetVisaOther, 'placeCityGetVisa'+ indexTab)
//     }
// }


