function loadFile(url, callback) {
    PizZipUtils.getBinaryContent(url, callback);
}

//виза

//визовая анкета
window.generateVisaApplication = function generate() {
    path = ('../Templates/виза/визовая анкета.docx')
    var zipDocs = new PizZip();
    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }

            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }
            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            for (let i =0; i<countTab();i++) {
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                let elem = tabs[i]
                let indexTab = parseInt(elem.id.match(/\d+/))

                // purpose
                let purposeG = ""
                let purposeR = ""
                let purposeU = ""
                let purposeS = "-"
                switch (document.getElementById('purpose' + indexTab).value) {
                    case "Учеба":
                        purposeU = "X"
                        purposeS = "Студент"
                        break
                    case "Краткосрочная учеба":
                        purposeU = "X"
                        purposeS = "Студент"
                        break
                    case "(НТС)":
                        purposeG = "X"
                        break
                    case "Трудовая деятельность":
                        purposeR = "X"
                        purposeS = "Преподаватель"
                        break
                }

                // Passport
                let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                    ? document.getElementById('series' + indexTab).value : ''
                let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                // gender
                let genM = ''
                let genW = ''
                switch (document.getElementById('gender' + indexTab).value) {
                    case "Мужской / Male":
                        genM = 'X'
                        break
                    case "Женский / Female":
                        genW = 'X'
                        break
                }

                // visa
                let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                    ? document.getElementById('seriesVisa' + indexTab).value : ''
                let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                    ? document.getElementById('idVisa' + indexTab).value : ''
                let identifierVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('identifierVisa' + indexTab).value)
                    ? document.getElementById('identifierVisa' + indexTab).value : ''
                let numInvVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('numInvVisa' + indexTab).value)
                    ? document.getElementById('numInvVisa' + indexTab).value : ''

                let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''


                // OVM
                let infHost1 = ''
                let infHost2 = ''
                let addressResidence = ''
                let numRoom = ''
                switch (document.getElementById('migrationAddress').value) {
                    case "Квартира":
                        if (document.getElementById('infHost' + indexTab).value.length>60) {
                            let totalLen = 0
                            let lenInfHost = document.getElementById('infHost' + indexTab).value.split(',')
                            for (let i = 0; i< lenInfHost.length; i++) {
                                totalLen = totalLen + lenInfHost[i].length
                                if (totalLen<60) {
                                    if (i==lenInfHost.length-1) {infHost1 = infHost1  + lenInfHost[i]}
                                    else {infHost1 = infHost1  + lenInfHost[i] + ', '}
                                }
                                else {
                                    if (i==lenInfHost.length-1) {infHost2 = infHost2  + lenInfHost[i]}
                                    else {infHost2 = infHost2  + lenInfHost[i] + ', '}
                                }
                            }
                        }
                        else {infHost1 = document.getElementById('infHost' + indexTab).value}
                        addressResidence = document.getElementById('addressResidence' + indexTab).value
                        break
                    default:
                        infHost1 = 'МПГУ, 7704077771, М. Пироговская д. 1, стр. 1.'
                        infHost2 = '8-499-245-03-10, mail@mpgu.su'
                        addressResidence = document.getElementById('migrationAddress').value
                        numRoom = ', комната № ' + document.getElementById('numRoom' + indexTab).value
                        break

                }

                doc.setData({


                    purposeG: purposeG,
                    purposeR: purposeR,
                    purposeU: purposeU,
                    lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                    lastNameEn: document.getElementById('lastNameEn' + indexTab).value.toUpperCase(),
                    firstNameRu: document.getElementById('firstNameRu' + indexTab).value.toUpperCase(),
                    firstNameEn: document.getElementById('firstNameEn' + indexTab).value.toUpperCase(),
                    patronymicRu: document.getElementById('patronymicRu' + indexTab).value.toUpperCase(),
                    dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                    placeStateBirth: document.getElementById('placeStateBirth' + indexTab).value.toUpperCase(),
                    genM: genM,
                    genW: genW,
                    grazd: document.getElementById('grazd' + indexTab).value,
                    // documentPerson: document.getElementById('documentPerson' + indexTab).text,
                    series: series,
                    idPassport: document.getElementById('idPassport' + indexTab).value,
                    dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                    validUntil: validUntil,

                    infHost1: infHost1,
                    infHost2: infHost2,
                    addressResidence: addressResidence,
                    numRoom: '',

                    homeAddress: document.getElementById('homeAddress' + indexTab).value,
                    purposeS: purposeS,
                    phone: document.getElementById('phone' + indexTab).value,
                    mail: document.getElementById('mail' + indexTab).value,
                    seriesVisa: seriesVisa,
                    idVisa: idVisa,
                    identifierVisa: identifierVisa,
                    dateOfIssueVisa: dateOfIssueVisa,
                    validUntilVisa: validUntilVisa,
                    numInvVisa: numInvVisa,
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),

                });



                try {
                    doc.render();
                }
                catch (error) {
                    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                    errorHandler(error);
                }
                var out = doc.getZip().generate();
                zipDocs.file("ВИЗОВАЯ АНКЕТА - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                    document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                    " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".docx"
                    , out, {base64: true}
                );
            } // end for


            let nameFile = ''
            if (countTab()==1) {
                nameFile = document.getElementById('nStud1').value
                    +" ВИЗОВАЯ АНКЕТА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text +".zip"
            }
            else {
                nameFile = document.getElementById('nStud1').value + '-'+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" ВИЗОВАЯ АНКЕТА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text +".zip"
            }
            var content = zipDocs.generate({ type: "blob" });
            saveAs(content,nameFile);
        });
};

//справка
window.generateVisaReference = function generate() {
    path = ('../Templates/виза/справка.docx')
    var zipDocs = new PizZip();
    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }

            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }
            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            for (let i =0; i<countTab();i++) {
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                let elem = tabs[i]
                let indexTab = parseInt(elem.id.match(/\d+/))

                // purpose
                let purpose = ""
                switch (document.getElementById('purpose' + indexTab).value) {
                    case "Учеба":
                        purpose = "студентом"
                        break
                    case "Краткосрочная учеба":
                        purpose = "студентом"
                        break
                    case "(НТС)":
                        purpose = "приглашенным гостем (НТС)"
                        break
                    case "Трудовая деятельность":
                        purpose = "преподавателем"
                        break
                }

                // faculty
                let faculty = ''
                switch (document.getElementById('faculty' + indexTab).value) {
                    case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                        faculty = 'Института изящных искусств: Факультета музыкального искусства'
                        break
                    case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                        faculty = 'Института изящных искусств: Художественно-графического факультета'
                        break
                    case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                        faculty = 'Института социально-гуманитарного образования'
                        break
                    case "Институт филологии / The Institute of Philology":
                        faculty = 'Института филологии'
                        break
                    case "Институт иностранных языков / The Institute of Foreign Languages":
                        faculty = 'Института иностранных языков'
                        break
                    case "Институт международного образования / The Institute of International Education":
                        faculty = 'Института международного образования'
                        break
                    case "Институт детства / The Institute of Childhood":
                        faculty = 'Института детства'
                        break
                    case "Институт биологии и химии / The Institute of Biology and Chemistry":
                        faculty = 'Института биологии и химии'
                        break
                    case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                        faculty = 'Института физики, технологии и информационных систем'
                        break
                    case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                        faculty = 'Института физической культуры, спорта и здоровья'
                        break
                    case "Географический факультет / The Institute of Geography":
                        faculty = 'Географического факультета'
                        break
                    case "Институт истории и политики / The Institute of History and Politics":
                        faculty = 'Института истории и политики'
                        break
                    case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                        faculty = 'Института математики и информатики'
                        break
                    case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                        faculty = 'Факультета дошкольной педагогики и психологии'
                        break
                    case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                        faculty = 'Института педагогики и психологии'
                        break
                    case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                        faculty = 'Института журналистики, коммуникаций и медиаобразования'
                        break
                    case "Институт развития цифрового образования / The Institute of Digital Education Development":
                        faculty = 'Института развития цифрового образования'
                        break

                }


                // Passport
                let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                    ? document.getElementById('series' + indexTab).value : ''
                let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''



                // OVM
                let ovmByRegion = ''
                switch (document.getElementById('ovmByRegion').value) {
                    case "Тропарево-Никулино":
                        ovmByRegion = 'ОМВД России по району Тропарево-Никулино г.Москвы'
                        break
                    case "Хамовники":
                        ovmByRegion = 'ОМВД России по району Хамовники г.Москвы'
                        break
                }

                // registration On
                let registrationOn = ''
                switch (document.getElementById('registrationOn').value) {
                    case "Круглов":
                        registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                        break
                    case "Морозова":
                        registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                        break
                    case "Орлова":
                        registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                        break
                }


                let dateUnt =  document.getElementById('dateUntil' + indexTab).value ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                doc.setData({

                    grazd: document.getElementById('grazd' + indexTab).value.toUpperCase(),
                    lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                    firstNameRu: document.getElementById('firstNameRu' + indexTab).value.toUpperCase(),
                    patronymicRu: document.getElementById('patronymicRu' + indexTab).value.toUpperCase(),
                    purpose: purpose,
                    faculty: faculty.toUpperCase(),
                    series: series,
                    idPassport: document.getElementById('idPassport' + indexTab).value,
                    dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                    validUntil: validUntil,
                    dateUntil: dateUnt,
                    ovmByRegion: ovmByRegion,
                    registrationOn: registrationOn,
                    nStud: document.getElementById('nStud' + indexTab).value,
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                });



                try {
                    doc.render();
                }
                catch (error) {
                    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                    errorHandler(error);
                }
                var out = doc.getZip().generate();
                zipDocs.file("СПРАВКА - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                    document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                    " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".docx"
                    , out, {base64: true}
                );
            } // end for


            let nameFile = ''
            if (countTab()==1) {
                nameFile = document.getElementById('nStud1').value
                    +" СПРАВКА - "  + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".zip"
            }
            else {
                nameFile = document.getElementById('nStud1').value + '-'+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" СПРАВКА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".zip"
            }
            var content = zipDocs.generate({ type: "blob" });
            saveAs(content,nameFile);
        });
};

//ходатайство ВИЗА ТРОПАРЕВО-НИКУЛИНО
window.generateVisaSolicitaionTroparevo = function generate() {
    path = ('../Templates/виза/ходатайство ТРОПАРЕВО-НИКУЛИНО.docx')
    var zipDocs = new PizZip();
    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }

            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }
            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            for (let i =0; i<countTab();i++) {
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                let elem = tabs[i]
                let indexTab = parseInt(elem.id.match(/\d+/))


                //dateUntil
                let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                // purpose
                let purpose = ''
                switch (document.getElementById('purpose' + indexTab).value) {
                    case "Учеба":
                        purpose = "обучением в МПГУ"
                        break
                    case "Краткосрочная учеба":
                        purpose = "обучением в МПГУ"
                        break
                    case "(НТС)":
                        purpose = "посещением МПГУ в качестве приглашенного гостя (НТС)"
                        break
                    case "Трудовая деятельность":
                        purpose = "преподавательской деятельностью в МПГУ"
                        break
                }

                // Passport
                let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                    ? document.getElementById('series' + indexTab).value : ''
                let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                // gender
                let genM = ''
                let genW = ''
                switch (document.getElementById('gender' + indexTab).value) {
                    case "Мужской / Male":
                        genM = 'X'
                        genW = ' '
                        break
                    case "Женский / Female":
                        genW = 'X'
                        genM = ' '
                        break
                }

                // registration On
                let registrationOn1 = ''
                let registrationOn2 = ''
                switch (document.getElementById('registrationOn').value) {
                    case "Круглов":
                        registrationOn1 = 'Начальник УМС                                                    Круглов В.В.'
                        registrationOn2 = 'Начальник УМС                                                                          В. В. Круглов'
                        break
                    case "Морозова":
                        registrationOn1 = 'Заместитель начальника УМС                Морозова О.А.'
                        registrationOn2 = 'Заместитель начальника УМС                                                Морозова О.А.'
                        break
                    case "Орлова":
                        registrationOn1 = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                        registrationOn2 = 'Начальник паспортно-визового отдела УМС                              Орлова С.В.'
                        break
                }

                // visa
                let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                    ? document.getElementById('seriesVisa' + indexTab).value : ''
                let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                    ? document.getElementById('idVisa' + indexTab).value : ''

                let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                // order
                let numOrder = document.getElementById('numOrder' + indexTab).value
                    ? document.getElementById('numOrder' + indexTab).value : ''
                let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''
                let orderUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString() : ''

                // contract
                let typeFunding = ''
                switch (document.getElementById('typeFunding' + indexTab).value) {
                    case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                        typeFunding = 'НАПРАВЛЕНИЕ №'
                        break
                    case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                        typeFunding = 'ДОГОВОР №'
                        break
                }
                let numContract = document.getElementById('numContract' + indexTab).value
                    ? document.getElementById('numContract' + indexTab).value : ''
                let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                    document.getElementById('contractFrom' + indexTab).value : ''



                doc.setData({
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    nStud: document.getElementById('nStud' + indexTab).value,
                    grazd: document.getElementById('grazd' + indexTab).value,
                    lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                    lastNameEn: document.getElementById('lastNameEn' + indexTab).value,
                    firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                    firstNameEn: document.getElementById('firstNameEn' + indexTab).value,
                    patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                    dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                    registrationOn1: registrationOn1,
                    registrationOn2: registrationOn2,
                    dateUntil: dateUntil,
                    genM: genM,
                    genW: genW,
                    purpose: purpose,

                    series: series,
                    idPassport: document.getElementById('idPassport' + indexTab).value,
                    dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                    validUntil: validUntil,

                    seriesVisa: seriesVisa,
                    idVisa: idVisa,
                    dateOfIssueVisa: dateOfIssueVisa,
                    validUntilVisa: validUntilVisa,


                    numOrder: numOrder,
                    orderFrom: orderFrom,
                    orderUntil: orderUntil,

                    typeFunding: typeFunding,
                    numContract: numContract,
                    contractFrom: contractFrom,

                });



                try {
                    doc.render();
                }
                catch (error) {
                    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                    errorHandler(error);
                }
                var out = doc.getZip().generate();
                zipDocs.file("ХОДАТАЙСТВО (ВИЗА) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                    document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                    " - " + "ОВМ ТРОПАРЕВО-НИКУЛИНО" + ".docx"
                    , out, {base64: true}
                );
            } // end for


            let nameFile = ''
            if (countTab()==1) {
                nameFile = document.getElementById('nStud1').value
                    +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ТРОПАРЕВО-НИКУЛИНО.zip"
            }
            else {
                nameFile = document.getElementById('nStud1').value + '-'+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ТРОПАРЕВО-НИКУЛИНО.zip"
            }
            var content = zipDocs.generate({ type: "blob" });
            saveAs(content,nameFile);
        });
};

//ходатайство ВИЗА ХАМОВНИКИ
window.generateVisaSolicitaionKhamovniki = function generate() {
    path = ('../Templates/виза/ходатайство ХАМОВНИКИ.docx')
    var zipDocs = new PizZip();
    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }

            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }
            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            for (let i =0; i<countTab();i++) {
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                let elem = tabs[i]
                let indexTab = parseInt(elem.id.match(/\d+/))


                //dateUntil
                let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                // purpose
                let purpose = ''
                let purposeS = ''
                switch (document.getElementById('purpose' + indexTab).value) {
                    case "Учеба":
                        purpose = "УЧЕБА"
                        purposeS = "Студент"
                        break
                    case "Краткосрочная учеба":
                        purpose = "КРАТКОСРОЧНАЯ УЧЕБА"
                        purposeS = "Студент"
                        break
                    case "(НТС)":
                        purpose = "НАУЧНО-ТЕХНИЧЕСКИЕ СВЯЗИ (НТС)"
                        purposeS = "НТС"
                        break
                    case "Трудовая деятельность":
                        purpose = "ТРУДОВАЯ ДЕЯТЕЛЬНОСТЬ"
                        purposeS = "Преподаватель"
                        break
                }

                // levelEducation
                let levelEducation = ''
                switch (document.getElementById('levelEducation' + indexTab).value) {
                    case "Подготовительный факультет (изучаю русский язык)/ The preparatory faculty":
                        levelEducation = 'подготовительный факультет'
                        break
                    case "бакалавриат/bachelor degree":
                        levelEducation = 'бакалавриат'
                        break
                    case "магистратура/master degree":
                        levelEducation = 'магистратура'
                        break
                    case "аспирантура/post-graduate studies":
                        levelEducation = 'аспирантура'
                        break
                }

                // course
                let course = ''
                switch (document.getElementById('course' + indexTab).value) {
                    case '1':
                        course = ', 1 курс,'
                        break
                    case '2':
                        course = ', 2 курс,'
                        break
                    case '3':
                        course = ', 3 курс,'
                        break
                    case '4':
                        course = ', 4 курс,'
                        break
                    case '5':
                        course = ', 5 курс,'
                        break
                }

                // norification
                let notificationFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('notificationFrom' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('notificationFrom' + indexTab).value).toLocaleDateString() : ''
                let notificationUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('notificationUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('notificationUntil' + indexTab).value).toLocaleDateString() : ''
                let issuedBy = document.getElementById('issuedBy' + indexTab).value != '' ? document.getElementById('issuedBy' + indexTab).value : ''

                // addressResidence
                let addressResidence = ''
                switch (document.getElementById('migrationAddress').value) {
                    case "Квартира":
                        addressResidence = document.getElementById('addressResidence' + indexTab).value
                        break
                    default:
                        addressResidence = document.getElementById('migrationAddress').value
                        break
                }

                // Passport
                let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                    ? document.getElementById('series' + indexTab).value : ''
                let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                // gender
                let gender = ''
                switch (document.getElementById('gender' + indexTab).value) {
                    case "Мужской / Male":
                        gender = 'м.'
                        break
                    case "Женский / Female":
                        gender = 'ж.'
                        break
                }

                // registration On
                let registrationOn = ''
                switch (document.getElementById('registrationOn').value) {
                    case "Круглов":
                        registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                        break
                    case "Морозова":
                        registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                        break
                    case "Орлова":
                        registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                        break
                }

                // visa
                let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                    ? document.getElementById('seriesVisa' + indexTab).value : ''
                let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                    ? document.getElementById('idVisa' + indexTab).value : ''

                let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                // order
                let numOrder = document.getElementById('numOrder' + indexTab).value
                    ? document.getElementById('numOrder' + indexTab).value : ''
                let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''

                // faculty
                let faculty = ''
                switch (document.getElementById('faculty' + indexTab).value) {
                    case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                        faculty = 'ИИИ:Музфак'
                        break
                    case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                        faculty = 'ИИИ: Худграф'
                        break
                    case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                        faculty = 'ИСГО'
                        break
                    case "Институт филологии / The Institute of Philology":
                        faculty = 'ИФ'
                        break
                    case "Институт иностранных языков / The Institute of Foreign Languages":
                        faculty = 'ИИЯ'
                        break
                    case "Институт международного образования / The Institute of International Education":
                        faculty = 'ИМО'
                        break
                    case "Институт детства / The Institute of Childhood":
                        faculty = 'ИД'
                        break
                    case "Институт биологии и химии / The Institute of Biology and Chemistry":
                        faculty = 'ИБХ'
                        break
                    case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                        faculty = 'ИФТИС'
                        break
                    case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                        faculty = 'ИФКСиЗ'
                        break
                    case "Географический факультет / The Institute of Geography":
                        faculty = 'Геофак'
                        break
                    case "Институт истории и политики / The Institute of History and Politics":
                        faculty = 'ИИП'
                        break
                    case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                        faculty = 'ИМИ'
                        break
                    case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                        faculty = 'Дош.фак.'
                        break
                    case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                        faculty = 'ИПП'
                        break
                    case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                        faculty = 'ИЖКиМ'
                        break
                    case "Институт развития цифрового образования / The Institute of Digital Education Development":
                        faculty = 'ИРЦО'
                        break
                }


                // contract
                let typeFundingDog1 = ""
                let typeFundingDog2 = ""
                let typeFundingNap1 = ""
                let typeFundingNap2 = ""
                switch (document.getElementById('typeFunding' + indexTab).value) {
                    case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                        typeFundingDog2 = "Договор"
                        typeFundingNap2 = "направление"
                        break
                    case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                        typeFundingDog1 = "Договор"
                        typeFundingNap1 = "направление"
                        break
                }
                let numContract = document.getElementById('numContract' + indexTab).value
                    ? document.getElementById('numContract' + indexTab).value : ''
                let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                    document.getElementById('contractFrom' + indexTab).value : ''



                doc.setData({
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    nStud: document.getElementById('nStud' + indexTab).value,
                    grazd: document.getElementById('grazd' + indexTab).value,
                    lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                    firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                    patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                    dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                    gender: gender,

                    registrationOn: registrationOn,
                    dateUntil: dateUntil,
                    purpose: purpose,
                    purposeS: purposeS,
                    levelEducation: levelEducation,
                    course: course,

                    series: series,
                    idPassport: document.getElementById('idPassport' + indexTab).value,
                    dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                    validUntil: validUntil,

                    seriesVisa: seriesVisa,
                    idVisa: idVisa,
                    dateOfIssueVisa: dateOfIssueVisa,
                    validUntilVisa: validUntilVisa,

                    seriesMigration: document.getElementById('seriesMigration' + indexTab).value,
                    idMigration: document.getElementById('idMigration' + indexTab).value,
                    dateArrivalMigration: new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString(),

                    notificationFrom: notificationFrom,
                    notificationUntil: notificationUntil,
                    issuedBy: issuedBy,

                    addressResidence: addressResidence,
                    faculty: faculty,
                    numOrder: numOrder,
                    orderFrom: orderFrom,
                    typeFundingDog1: typeFundingDog1,
                    typeFundingDog2: typeFundingDog2,
                    typeFundingNap1: typeFundingNap1,
                    typeFundingNap2: typeFundingNap2,

                    numContract: numContract,
                    contractFrom: contractFrom,

                });



                try {
                    doc.render();
                }
                catch (error) {
                    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                    errorHandler(error);
                }
                var out = doc.getZip().generate();
                zipDocs.file("ХОДАТАЙСТВО (ВИЗА) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                    document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                    " - " + "ОВМ ХАМОВНИКИ" + ".docx"
                    , out, {base64: true}
                );
            } // end for


            let nameFile = ''
            if (countTab()==1) {
                nameFile = document.getElementById('nStud1').value
                    +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ХАМОВНИКИ.zip"
            }
            else {
                nameFile = document.getElementById('nStud1').value + '-'+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ХАМОВНИКИ.zip"
            }
            var content = zipDocs.generate({ type: "blob" });
            saveAs(content,nameFile);
        });
};

// выбор функции для ходатайства, относительно выбора ОВМ
function generateVisaSolic(x) {
    if (x=='Тропарево-Никулино') {
        generateVisaSolicitaionTroparevo()
    }
    else if (x=='Хамовники') {
        generateVisaSolicitaionKhamovniki()
    }
}

//опись ВИЗА
window.generateInventoryVisa = function generate() {
    path = ('../Templates/виза/опись виза.docx')

    let students = []

    // registration On
    let registrationOn = ''
    switch (document.getElementById('registrationOn').value) {
        case "Круглов":
            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
            break
        case "Морозова":
            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
            break
        case "Орлова":
            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
            break
    }

    // ovmByRegion
    let ovmByRegion = ''
    switch (document.getElementById('ovmByRegion').value) {
        case 'Тропарево-Никулино':
            ovmByRegion = 'ОМВД России по району Тропарево-Никулино г. Москвы'
            break
        case 'Хамовники':
            ovmByRegion = 'ОМВД России по району Хамовники г. Москвы'
            break
        case 'Алексеевский':
            ovmByRegion = 'ОМВД России по Алексеевскому району г.Москвы'
            break
        case 'Войковский':
            ovmByRegion = 'ОВМ МВД России по району Войковский г. Москвы'
            break
        case 'МУ МВД РФ Люберецкое':
            ovmByRegion = 'ОВМ МУ МВД РФ "Люберецкое"'
            break
    }

    // nStud
    let nStud1 = document.getElementById('nStud1').value
    let tir = ''
    let nStud2 = ''
    if (countTab()>1) {
        nStud2 = document.getElementById('nStud' + (lastTab()-1)).value
        tir = '-'
    }

    for (let i =0; i<countTab();i++) {
        let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
        let elem = tabs[i]
        let indexTab = parseInt(elem.id.match(/\d+/))



        let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
            ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''

        let dateUntil = document.getElementById('dateUntil' + indexTab).value != '-' ?
            new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

        //students
        students.push({
            nStud: document.getElementById('nStud' + indexTab).value,
            lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
            firstNameRu: document.getElementById('firstNameRu'+indexTab).value.toUpperCase(),
            patronymicRu: document.getElementById('patronymicRu'+indexTab).value.toUpperCase(),
            dateOfIssueVisa: dateOfIssueVisa,
            dateUntil: dateUntil,
            grazd: document.getElementById('grazd' + indexTab).value,
            phone: document.getElementById('phone' + indexTab).value,
            mail: document.getElementById('mail' + indexTab).value,
        })
    }


    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }
            var zip = new PizZip(content);
            var doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            doc.render({
                dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                nStud1: nStud1,
                nStud2: nStud2,
                tir: tir,
                ovmByRegion: ovmByRegion,
                'students': students,
                registrationOn: registrationOn,
            })


            var out = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                // compression: DEFLATE adds a compression step.
                // For a 50MB output document, expect 500ms additional CPU time
                compression: "DEFLATE",
            });



            let nameFile = ''
            if (countTab()==1) {
                nameFile = "ОПИСЬ (ВИЗА) - студент " +
                    document.getElementById('nStud1').value
                    +" - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                    +" - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
            }
            else {
                nameFile = "ОПИСЬ (ВИЗА) - студенты " +
                    document.getElementById('nStud1').value
                    +"-"+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                    +" - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
            }


            // Output the document using Data-URI
            saveAs(out, nameFile);
        }
    );
};





//регистрация

//ходатайство РЕГИСТРАЦИЯ
window.generateRegSolicitaion = function generate() {
    path = ("")
    switch (document.getElementById('ovmByRegion').value) {
        case "Алексеевский":
            path = ('../Templates/регистрация/ходатайство АЛЕКСЕЕВСКИЙ.docx')
            break
        case "Войковский":
            path = ('../Templates/регистрация/ходатайство ВОЙКОВСКИЙ.docx')
            break
        case "МУ МВД РФ Люберецкое":
            path = ('../Templates/регистрация/ходатайство МУ МВД РФ ЛЮБЕРЕЦКОЕ.docx')
            break
        case "Тропарево-Никулино":
            path = ('../Templates/регистрация/ходатайство ТРОПАРЕВО-НИКУЛИНО.docx')
            break
    }

    var zipDocs = new PizZip();
    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }

            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }
            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            for (let i =0; i<countTab();i++) {
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                let elem = tabs[i]
                let indexTab = parseInt(elem.id.match(/\d+/))


                //dateUntil
                let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                // purpose
                let purpose = ''
                let purposeS = ''
                switch (document.getElementById('purpose' + indexTab).value) {
                    case "Учеба":
                        purpose = "УЧЕБА"
                        purposeS = "Студент"
                        break
                    case "Краткосрочная учеба":
                        purpose = "КРАТКОСРОЧНАЯ УЧЕБА"
                        purposeS = "Студент"
                        break
                    case "(НТС)":
                        purpose = "НАУЧНО-ТЕХНИЧЕСКИЕ СВЯЗИ (НТС)"
                        purposeS = "НТС"
                        break
                    case "Трудовая деятельность":
                        purpose = "ТРУДОВАЯ ДЕЯТЕЛЬНОСТЬ"
                        purposeS = "Преподаватель"
                        break
                }

                // levelEducation
                let levelEducation = ''
                switch (document.getElementById('levelEducation' + indexTab).value) {
                    case "Подготовительный факультет (изучаю русский язык)/ The preparatory faculty":
                        levelEducation = 'подготовительный факультет'
                        break
                    case "бакалавриат/bachelor degree":
                        levelEducation = 'бакалавриат'
                        break
                    case "магистратура/master degree":
                        levelEducation = 'магистратура'
                        break
                    case "аспирантура/post-graduate studies":
                        levelEducation = 'аспирантура'
                        break
                }

                // course
                let course = ''
                switch (document.getElementById('course' + indexTab).value) {
                    case '1':
                        course = ', 1 курс,'
                        break
                    case '2':
                        course = ', 2 курс,'
                        break
                    case '3':
                        course = ', 3 курс,'
                        break
                    case '4':
                        course = ', 4 курс,'
                        break
                    case '5':
                        course = ', 5 курс,'
                        break
                }

                // addressResidence
                let addressResidence = ''
                switch (document.getElementById('migrationAddress').value) {
                    case "Квартира":
                        addressResidence = document.getElementById('addressResidence' + indexTab).value
                        break
                    default:
                        addressResidence = document.getElementById('migrationAddress').value
                        break
                }

                // Passport
                let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                    ? document.getElementById('series' + indexTab).value : ''
                let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                // gender
                let gender = ''
                switch (document.getElementById('gender' + indexTab).value) {
                    case "Мужской / Male":
                        gender = 'м.'
                        break
                    case "Женский / Female":
                        gender = 'ж.'
                        break
                }

                // registration On
                let registrationOn = ''
                switch (document.getElementById('registrationOn').value) {
                    case "Круглов":
                        registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                        break
                    case "Морозова":
                        registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                        break
                    case "Орлова":
                        registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                        break
                }

                // visa
                let typeVisa = ''
                switch (document.getElementById('typeVisa' + indexTab).value) {
                    case "ВИЗА":
                        typeVisa = 'Виза'
                        break
                    case "(ВНЖ) ВИД НА ЖИТЕЛЬСТВО РФ":
                        typeVisa = 'ВНЖ'
                        break
                    case "(РВП) РАЗРЕШЕНИЕ НА ВРЕМЕННОЕ ПРОЖИВАНИЕ РФ":
                        typeVisa = 'РВП'
                        break

                }
                let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                    ? document.getElementById('seriesVisa' + indexTab).value : ''
                let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                    ? document.getElementById('idVisa' + indexTab).value : ''

                let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                // order
                let numOrder = document.getElementById('numOrder' + indexTab).value
                    ? document.getElementById('numOrder' + indexTab).value : ''
                let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''
                let orderUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString() : ''

                // faculty
                let faculty = ''
                switch (document.getElementById('faculty' + indexTab).value) {
                    case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                        faculty = 'ИИИ:Музфак'
                        break
                    case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                        faculty = 'ИИИ: Худграф'
                        break
                    case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                        faculty = 'ИСГО'
                        break
                    case "Институт филологии / The Institute of Philology":
                        faculty = 'ИФ'
                        break
                    case "Институт иностранных языков / The Institute of Foreign Languages":
                        faculty = 'ИИЯ'
                        break
                    case "Институт международного образования / The Institute of International Education":
                        faculty = 'ИМО'
                        break
                    case "Институт детства / The Institute of Childhood":
                        faculty = 'ИД'
                        break
                    case "Институт биологии и химии / The Institute of Biology and Chemistry":
                        faculty = 'ИБХ'
                        break
                    case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                        faculty = 'ИФТИС'
                        break
                    case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                        faculty = 'ИФКСиЗ'
                        break
                    case "Географический факультет / The Institute of Geography":
                        faculty = 'Геофак'
                        break
                    case "Институт истории и политики / The Institute of History and Politics":
                        faculty = 'ИИП'
                        break
                    case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                        faculty = 'ИМИ'
                        break
                    case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                        faculty = 'Дош.фак.'
                        break
                    case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                        faculty = 'ИПП'
                        break
                    case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                        faculty = 'ИЖКиМ'
                        break
                    case "Институт развития цифрового образования / The Institute of Digital Education Development":
                        faculty = 'ИРЦО'
                        break
                }


                // contract
                let typeFundingDog1 = ""
                let typeFundingDog2 = ""
                let typeFundingNap1 = ""
                let typeFundingNap2 = ""
                switch (document.getElementById('typeFunding' + indexTab).value) {
                    case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                        typeFundingDog2 = "Договор"
                        typeFundingNap2 = "направление"
                        break
                    case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                        typeFundingDog1 = "Договор"
                        typeFundingNap1 = "направление"
                        break
                }
                let numContract = document.getElementById('numContract' + indexTab).value
                    ? document.getElementById('numContract' + indexTab).value : ''
                let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                    document.getElementById('contractFrom' + indexTab).value : ''


                let numRoom = document.getElementById('numRoom' + indexTab).value != "" ? ', комната № ' + document.getElementById('numRoom' + indexTab).value : ''

                let numRental = document.getElementById('numRental' + indexTab).value != "-" ? document.getElementById('numRental' + indexTab).value : ''


                doc.setData({
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    nStud: document.getElementById('nStud' + indexTab).value,
                    grazd: document.getElementById('grazd' + indexTab).value,
                    lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                    firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                    patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                    dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                    gender: gender,

                    registrationOn: registrationOn,
                    dateUntil: dateUntil,
                    purpose: purpose,
                    purposeS: purposeS,
                    levelEducation: levelEducation,
                    course: course,

                    series: series,
                    idPassport: document.getElementById('idPassport' + indexTab).value,
                    dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                    validUntil: validUntil,

                    typeVisa: typeVisa,
                    seriesVisa: seriesVisa,
                    idVisa: idVisa,
                    dateOfIssueVisa: dateOfIssueVisa,
                    validUntilVisa: validUntilVisa,

                    seriesMigration: document.getElementById('seriesMigration' + indexTab).value,
                    idMigration: document.getElementById('idMigration' + indexTab).value,
                    dateArrivalMigration: new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString(),



                    migrationAddress: document.getElementById('migrationAddress').value,
                    numRoom: '',
                    faculty: faculty,
                    numOrder: numOrder,
                    orderFrom: orderFrom,
                    orderUntil: orderUntil,

                    typeFundingDog1: typeFundingDog1,
                    typeFundingDog2: typeFundingDog2,
                    typeFundingNap1: typeFundingNap1,
                    typeFundingNap2: typeFundingNap2,

                    numContract: numContract,
                    contractFrom: contractFrom,

                    numRental: '',

                });



                try {
                    doc.render();
                }
                catch (error) {
                    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                    errorHandler(error);
                }
                var out = doc.getZip().generate();
                zipDocs.file("ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                    document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                    " - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
                    , out, {base64: true}
                );
            } // end for


            let nameFile = ''
            if (countTab()==1) {
                nameFile = document.getElementById('nStud1').value
                    +" ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    +".zip"
            }
            else {
                nameFile = document.getElementById('nStud1').value + '-'+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".zip"
            }
            var content = zipDocs.generate({ type: "blob" });
            saveAs(content,nameFile);
        });
};

//опись РЕГИСТРАЦИЯ
window.generateInventoryReg = function generate() {
    path = ('../Templates/регистрация/опись регистрация.docx')

    let students = []

    // registration On
    let registrationOn = ''
    switch (document.getElementById('registrationOn').value) {
        case "Круглов":
            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
            break
        case "Морозова":
            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
            break
        case "Орлова":
            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
            break
    }

    // ovmByRegion
    let ovmByRegion = ''
    switch (document.getElementById('ovmByRegion').value) {
        case 'Тропарево-Никулино':
            ovmByRegion = 'ОМВД России по району Тропарево-Никулино г. Москвы'
            break
        case 'Хамовники':
            ovmByRegion = 'ОМВД России по району Хамовники г. Москвы'
            break
        case 'Алексеевский':
            ovmByRegion = 'ОМВД России по Алексеевскому району г.Москвы'
            break
        case 'Войковский':
            ovmByRegion = 'ОВМ МВД России по району Войковский г. Москвы'
            break
        case 'МУ МВД РФ Люберецкое':
            ovmByRegion = 'ОВМ МУ МВД РФ "Люберецкое"'
            break
    }

    // nStud
    let nStud1 = document.getElementById('nStud1').value
    let tir = ''
    let nStud2 = ''
    if (countTab()>1) {
        nStud2 = document.getElementById('nStud' + (lastTab()-1)).value
        tir = '-'
    }

    for (let i =0; i<countTab();i++) {
        let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
        let elem = tabs[i]
        let indexTab = parseInt(elem.id.match(/\d+/))



        let dateUntil = document.getElementById('dateUntil' + indexTab).value != '-' ?
            new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

        //students
        students.push({
            nStud: document.getElementById('nStud' + indexTab).value,
            lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
            firstNameRu: document.getElementById('firstNameRu'+indexTab).value.toUpperCase(),
            patronymicRu: document.getElementById('patronymicRu'+indexTab).value.toUpperCase(),
            dateInOv: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
            dateUntil: dateUntil,
            grazd: document.getElementById('grazd' + indexTab).value,
            phone: document.getElementById('phone' + indexTab).value,
            mail: document.getElementById('mail' + indexTab).value,
        })
    }


    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }
            var zip = new PizZip(content);
            var doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            doc.render({
                dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                nStud1: nStud1,
                nStud2: nStud2,
                tir: tir,
                ovmByRegion: ovmByRegion,
                'students': students,
                registrationOn: registrationOn,
            })


            var out = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                // compression: DEFLATE adds a compression step.
                // For a 50MB output document, expect 500ms additional CPU time
                compression: "DEFLATE",
            });



            let nameFile = ''
            if (countTab()==1) {
                nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ) - студент " +
                    document.getElementById('nStud1').value
                    +" - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                    +" - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
            }
            else {
                nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ) - студенты " +
                    document.getElementById('nStud1').value
                    +"-"+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                    +" - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
            }


            // Output the document using Data-URI
            saveAs(out, nameFile);
        }
    );
};

//уведомление РЕГИСТРАЦИЯ
window.generateRegNotif = function generate() {
    let ovmRg = document.getElementById('ovmByRegion').value
    let rgOn = document.getElementById('registrationOn').value
    path = (`../Templates/регистрация/уведомление ${ovmRg} ${rgOn}.docx`)

    var zipDocs = new PizZip();
    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }

            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }
            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            for (let i =0; i<countTab();i++) {
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                let elem = tabs[i]
                let indexTab = parseInt(elem.id.match(/\d+/))

                // lastName {lN1-27}
                let lN = document.getElementById('lastNameRu'+indexTab).value.toUpperCase().split('')
                let lN1 = (lN[0]) ? (lN[0]) : ''
                let lN2 = (lN[1]) ? (lN[1]) : ''
                let lN3 = (lN[2]) ? (lN[2]) : ''
                let lN4 = (lN[3]) ? (lN[3]) : ''
                let lN5 = (lN[4]) ? (lN[4]) : ''
                let lN6 = (lN[5]) ? (lN[5]) : ''
                let lN7 = (lN[6]) ? (lN[6]) : ''
                let lN8 = (lN[7]) ? (lN[7]) : ''
                let lN9 = (lN[8]) ? (lN[8]) : ''
                let lN10 = (lN[9]) ? (lN[9]) : ''
                let lN11 = (lN[10]) ? (lN[10]) : ''
                let lN12 = (lN[11]) ? (lN[11]) : ''
                let lN13 = (lN[12]) ? (lN[12]) : ''
                let lN14 = (lN[13]) ? (lN[13]) : ''
                let lN15 = (lN[14]) ? (lN[14]) : ''
                let lN16 = (lN[15]) ? (lN[15]) : ''
                let lN17 = (lN[16]) ? (lN[16]) : ''
                let lN18 = (lN[17]) ? (lN[17]) : ''
                let lN19 = (lN[18]) ? (lN[18]) : ''
                let lN20 = (lN[19]) ? (lN[19]) : ''
                let lN21 = (lN[20]) ? (lN[20]) : ''
                let lN22 = (lN[21]) ? (lN[21]) : ''
                let lN23 = (lN[22]) ? (lN[22]) : ''
                let lN24 = (lN[23]) ? (lN[23]) : ''
                let lN25 = (lN[24]) ? (lN[24]) : ''
                let lN26 = (lN[25]) ? (lN[25]) : ''
                let lN27 = (lN[26]) ? (lN[26]) : ''

                // firstName {fN1-f27}
                let fN = document.getElementById('firstNameRu' + indexTab).value.toUpperCase().split('')
                let fN1 = (fN[0]) ? (fN[0]) : ''
                let fN2 = (fN[1]) ? (fN[1]) : ''
                let fN3 = (fN[2]) ? (fN[2]) : ''
                let fN4 = (fN[3]) ? (fN[3]) : ''
                let fN5 = (fN[4]) ? (fN[4]) : ''
                let fN6 = (fN[5]) ? (fN[5]) : ''
                let fN7 = (fN[6]) ? (fN[6]) : ''
                let fN8 = (fN[7]) ? (fN[7]) : ''
                let fN9 = (fN[8]) ? (fN[8]) : ''
                let fN10 = (fN[9]) ? (fN[9]) : ''
                let fN11 = (fN[10]) ? (fN[10]) : ''
                let fN12 = (fN[11]) ? (fN[11]) : ''
                let fN13 = (fN[12]) ? (fN[12]) : ''
                let fN14 = (fN[13]) ? (fN[13]) : ''
                let fN15 = (fN[14]) ? (fN[14]) : ''
                let fN16 = (fN[15]) ? (fN[15]) : ''
                let fN17 = (fN[16]) ? (fN[16]) : ''
                let fN18 = (fN[17]) ? (fN[17]) : ''
                let fN19 = (fN[18]) ? (fN[18]) : ''
                let fN20 = (fN[19]) ? (fN[19]) : ''
                let fN21 = (fN[20]) ? (fN[20]) : ''
                let fN22 = (fN[21]) ? (fN[21]) : ''
                let fN23 = (fN[22]) ? (fN[22]) : ''
                let fN24 = (fN[23]) ? (fN[23]) : ''
                let fN25 = (fN[24]) ? (fN[24]) : ''
                let fN26 = (fN[25]) ? (fN[25]) : ''
                let fN27 = (fN[26]) ? (fN[26]) : ''

                // patronymic {patr1-24}
                let patr = document.getElementById('patronymicRu' + indexTab).value.toUpperCase().split('')
                let patr1 = (patr[0]) ? (patr[0]) : ''
                let patr2 = (patr[1]) ? (patr[1]) : ''
                let patr3 = (patr[2]) ? (patr[2]) : ''
                let patr4 = (patr[3]) ? (patr[3]) : ''
                let patr5 = (patr[4]) ? (patr[4]) : ''
                let patr6 = (patr[5]) ? (patr[5]) : ''
                let patr7 = (patr[6]) ? (patr[6]) : ''
                let patr8 = (patr[7]) ? (patr[7]) : ''
                let patr9 = (patr[8]) ? (patr[8]) : ''
                let patr10 = (patr[9]) ? (patr[9]) : ''
                let patr11 = (patr[10]) ? (patr[10]) : ''
                let patr12 = (patr[11]) ? (patr[11]) : ''
                let patr13 = (patr[12]) ? (patr[12]) : ''
                let patr14 = (patr[13]) ? (patr[13]) : ''
                let patr15 = (patr[14]) ? (patr[14]) : ''
                let patr16 = (patr[15]) ? (patr[15]) : ''
                let patr17 = (patr[16]) ? (patr[16]) : ''
                let patr18 = (patr[17]) ? (patr[17]) : ''
                let patr19 = (patr[18]) ? (patr[18]) : ''
                let patr20 = (patr[19]) ? (patr[19]) : ''
                let patr21 = (patr[20]) ? (patr[20]) : ''
                let patr22 = (patr[21]) ? (patr[21]) : ''
                let patr23 = (patr[22]) ? (patr[22]) : ''
                let patr24 = (patr[23]) ? (patr[23]) : ''

                // grazd {grazd1-25}
                let grazd = document.getElementById('grazd' + indexTab).value.toUpperCase().split('')
                let grazd1 = (grazd[0]) ? (grazd[0]) : ''
                let grazd2 = (grazd[1]) ? (grazd[1]) : ''
                let grazd3 = (grazd[2]) ? (grazd[2]) : ''
                let grazd4 = (grazd[3]) ? (grazd[3]) : ''
                let grazd5 = (grazd[4]) ? (grazd[4]) : ''
                let grazd6 = (grazd[5]) ? (grazd[5]) : ''
                let grazd7 = (grazd[6]) ? (grazd[6]) : ''
                let grazd8 = (grazd[7]) ? (grazd[7]) : ''
                let grazd9 = (grazd[8]) ? (grazd[8]) : ''
                let grazd10 = (grazd[9]) ? (grazd[9]) : ''
                let grazd11 = (grazd[10]) ? (grazd[10]) : ''
                let grazd12 = (grazd[11]) ? (grazd[11]) : ''
                let grazd13 = (grazd[12]) ? (grazd[12]) : ''
                let grazd14 = (grazd[13]) ? (grazd[13]) : ''
                let grazd15 = (grazd[14]) ? (grazd[14]) : ''
                let grazd16 = (grazd[15]) ? (grazd[15]) : ''
                let grazd17 = (grazd[16]) ? (grazd[16]) : ''
                let grazd18 = (grazd[17]) ? (grazd[17]) : ''
                let grazd19 = (grazd[18]) ? (grazd[18]) : ''
                let grazd20 = (grazd[19]) ? (grazd[19]) : ''
                let grazd21 = (grazd[20]) ? (grazd[20]) : ''
                let grazd22 = (grazd[21]) ? (grazd[21]) : ''
                let grazd23 = (grazd[22]) ? (grazd[22]) : ''
                let grazd24 = (grazd[23]) ? (grazd[23]) : ''
                let grazd25 = (grazd[24]) ? (grazd[24]) : ''

                // dateOfBirth {dOB1-8}
                let dOB = new Date(document.getElementById('dateOfBirth'+indexTab).value).toLocaleDateString().split('.')
                let dOB1 = (dOB) ? (dOB[0][0]) : ''
                let dOB2 = (dOB) ? (dOB[0][1]) : ''
                let dOB3 = (dOB) ? (dOB[1][0]) : ''
                let dOB4 = (dOB) ? (dOB[1][1]) : ''
                let dOB5 = (dOB) ? (dOB[2][0]) : ''
                let dOB6 = (dOB) ? (dOB[2][1]) : ''
                let dOB7 = (dOB) ? (dOB[2][2]) : ''
                let dOB8 = (dOB) ? (dOB[2][3]) : ''

                // gender
                let genM = ''
                let genW = ''
                switch (document.getElementById('gender' + indexTab).value) {
                    case "Мужской / Male":
                        genM = 'X'
                        break
                    case "Женский / Female":
                        genW = 'X'
                        break
                }

                // placeStateBirth {pSBS1-24} для страны И {pSBG1-24} для города
                let pSBS = document.getElementById('placeStateBirth'+indexTab).value.toUpperCase().split(' ')
                
                let pSBS1 = (pSBS[0][0]) ? (pSBS[0][0]) : ''
                let pSBS2 = (pSBS[0][1]) ? (pSBS[0][1]) : ''
                let pSBS3 = (pSBS[0][2]) ? (pSBS[0][2]) : ''
                let pSBS4 = (pSBS[0][3]) ? (pSBS[0][3]) : ''
                let pSBS5 = (pSBS[0][4]) ? (pSBS[0][4]) : ''
                let pSBS6 = (pSBS[0][5]) ? (pSBS[0][5]) : ''
                let pSBS7 = (pSBS[0][6]) ? (pSBS[0][6]) : ''
                let pSBS8 = (pSBS[0][7]) ? (pSBS[0][7]) : ''
                let pSBS9 = (pSBS[0][8]) ? (pSBS[0][8]) : ''
                let pSBS10 = (pSBS[0][9]) ? (pSBS[0][9]) : ''
                let pSBS11 = (pSBS[0][10]) ? (pSBS[0][10]) : ''
                let pSBS12 = (pSBS[0][11]) ? (pSBS[0][11]) : ''
                let pSBS13 = (pSBS[0][12]) ? (pSBS[0][12]) : ''
                let pSBS14 = (pSBS[0][13]) ? (pSBS[0][13]) : ''
                let pSBS15 = (pSBS[0][14]) ? (pSBS[0][14]) : ''
                let pSBS16 = (pSBS[0][15]) ? (pSBS[0][15]) : ''
                let pSBS17 = (pSBS[0][16]) ? (pSBS[0][16]) : ''
                let pSBS18 = (pSBS[0][17]) ? (pSBS[0][17]) : ''
                let pSBS19 = (pSBS[0][18]) ? (pSBS[0][18]) : ''
                let pSBS20 = (pSBS[0][19]) ? (pSBS[0][19]) : ''
                let pSBS21 = (pSBS[0][20]) ? (pSBS[0][20]) : ''
                let pSBS22 = (pSBS[0][21]) ? (pSBS[0][21]) : ''
                let pSBS23 = (pSBS[0][22]) ? (pSBS[0][22]) : ''
                let pSBS24 = (pSBS[0][23]) ? (pSBS[0][23]) : ''
                let pSBG1 = ""
                let pSBG2 = ""
                let pSBG3 = ""
                let pSBG4 = ""
                let pSBG5 = ""
                let pSBG6 = ""
                let pSBG7 = ""
                let pSBG8 = ""
                let pSBG9 = ""
                let pSBG10 = ""
                let pSBG11 = ""
                let pSBG12 = ""
                let pSBG13 = ""
                let pSBG14 = ""
                let pSBG15 = ""
                let pSBG16 = ""
                let pSBG17 = ""
                let pSBG18 = ""
                let pSBG19 = ""
                let pSBG20 = ""
                let pSBG21 = ""
                let pSBG22 = ""
                let pSBG23 = ""
                let pSBG24 = ""
                if (pSBS.length>1) {
                    pSBG1 = (pSBS[1][0]) ? (pSBS[1][0]) : ''
                    pSBG2 = (pSBS[1][1]) ? (pSBS[1][1]) : ''
                    pSBG3 = (pSBS[1][2]) ? (pSBS[1][2]) : ''
                    pSBG4 = (pSBS[1][3]) ? (pSBS[1][3]) : ''
                    pSBG5 = (pSBS[1][4]) ? (pSBS[1][4]) : ''
                    pSBG6 = (pSBS[1][5]) ? (pSBS[1][5]) : ''
                    pSBG7 = (pSBS[1][6]) ? (pSBS[1][6]) : ''
                    pSBG8 = (pSBS[1][7]) ? (pSBS[1][7]) : ''
                    pSBG9 = (pSBS[1][8]) ? (pSBS[1][8]) : ''
                    pSBG10 = (pSBS[1][9]) ? (pSBS[1][9]) : ''
                    pSBG11 = (pSBS[1][10]) ? (pSBS[1][10]) : ''
                    pSBG12 = (pSBS[1][11]) ? (pSBS[1][11]) : ''
                    pSBG13 = (pSBS[1][12]) ? (pSBS[1][12]) : ''
                    pSBG14 = (pSBS[1][13]) ? (pSBS[1][13]) : ''
                    pSBG15 = (pSBS[1][14]) ? (pSBS[1][14]) : ''
                    pSBG16 = (pSBS[1][15]) ? (pSBS[1][15]) : ''
                    pSBG17 = (pSBS[1][16]) ? (pSBS[1][16]) : ''
                    pSBG18 = (pSBS[1][17]) ? (pSBS[1][17]) : ''
                    pSBG19 = (pSBS[1][18]) ? (pSBS[1][18]) : ''
                    pSBG20 = (pSBS[1][19]) ? (pSBS[1][19]) : ''
                    pSBG21 = (pSBS[1][20]) ? (pSBS[1][20]) : ''
                    pSBG22 = (pSBS[1][21]) ? (pSBS[1][21]) : ''
                    pSBG23 = (pSBS[1][22]) ? (pSBS[1][22]) : ''
                    pSBG24 = (pSBS[1][23]) ? (pSBS[1][23]) : ''
                }

                // series {sP1-4}
                let sP = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                    ? document.getElementById('series' + indexTab).value.toUpperCase() : ''
                let sP1 = ''
                let sP2 = ''
                let sP3 = ''
                let sP4 = ''
                if (sP) {
                    switch (sP.length) {
                        case 1:
                            sP4 = sP[0]
                            break
                        case 2:
                            sP3 = sP[0]
                            sP4 = sP[1]
                            break
                        case 3:
                            sP2 = sP[0]
                            sP3 = sP[1]
                            sP4 = sP[2]
                            break
                        case 4:
                            sP1 = sP[0]
                            sP2 = sP[1]
                            sP3 = sP[2]
                            sP4 = sP[3]
                    }
                }


                // idPassport {iP1-10}
                let idP = document.getElementById('idPassport'+indexTab).value.toUpperCase()
                let idP1 = (idP[0]) ? (idP[0]) : ''
                let idP2 = (idP[1]) ? (idP[1]) : ''
                let idP3 = (idP[2]) ? (idP[2]) : ''
                let idP4 = (idP[3]) ? (idP[3]) : ''
                let idP5 = (idP[4]) ? (idP[4]) : ''
                let idP6 = (idP[5]) ? (idP[5]) : ''
                let idP7 = (idP[6]) ? (idP[6]) : ''
                let idP8 = (idP[7]) ? (idP[7]) : ''
                let idP9 = (idP[8]) ? (idP[8]) : ''
                let idP10 = (idP[9]) ? (idP[9]) : ''

                // dateOfIssue {dOI1-8}
                let dOI = new Date(document.getElementById('dateOfIssue'+indexTab).value).toLocaleDateString().split('.')
                let dOI1 = (dOI) ? (dOI[0][0]) : ''
                let dOI2 = (dOI) ? (dOI[0][1]) : ''
                let dOI3 = (dOI) ? (dOI[1][0]) : ''
                let dOI4 = (dOI) ? (dOI[1][1]) : ''
                let dOI5 = (dOI) ? (dOI[2][0]) : ''
                let dOI6 = (dOI) ? (dOI[2][1]) : ''
                let dOI7 = (dOI) ? (dOI[2][2]) : ''
                let dOI8 = (dOI) ? (dOI[2][3]) : ''


                // validUntil {vU1-8}
                let vU = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString().split(".") : ''
                let vU1 = (vU) ? (vU[0][0]) : ''
                let vU2 = (vU) ? (vU[0][1]) : ''
                let vU3 = (vU) ? (vU[1][0]) : ''
                let vU4 = (vU) ? (vU[1][1]) : ''
                let vU5 = (vU) ? (vU[2][0]) : ''
                let vU6 = (vU) ? (vU[2][1]) : ''
                let vU7 = (vU) ? (vU[2][2]) : ''
                let vU8 = (vU) ? (vU[2][3]) : ''


                // typeVisa
                let tVV = ''
                let tVJ = ''
                let tVP = ''
                switch (document.getElementById('typeVisa' + indexTab).value) {
                    case "ВИЗА":
                        tVV = 'X'
                        break
                    case "(ВНЖ) ВИД НА ЖИТЕЛЬСТВО РФ":
                        tVJ = 'X'
                        break
                    case "(РВП) РАЗРЕШЕНИЕ НА ВРЕМЕННОЕ ПРОЖИВАНИЕ РФ":
                        tVP = 'X'
                        break
                }

                // seriesVisa {sV1-4}
                let sV = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                    ? document.getElementById('seriesVisa' + indexTab).value.toUpperCase() : ''
                let sV1 = ''
                let sV2 = ''
                let sV3 = ''
                let sV4 = ''
                if (sV) {
                    switch (sV.length) {
                        case 1:
                            sV4 = sV[0]
                            break
                        case 2:
                            sV3 = sV[0]
                            sV4 = sV[1]
                            break
                        case 3:
                            sV2 = sV[0]
                            sV3 = sV[1]
                            sV4 = sV[2]
                            break
                        case 4:
                            sV1 = sV[0]
                            sV2 = sV[1]
                            sV3 = sV[2]
                            sV4 = sV[3]
                    }
                }

                // idVisa {idV1-15}
                let idV = document.getElementById('idVisa' + indexTab).value ? document.getElementById('idVisa' + indexTab).value.toUpperCase() : ''
                let idV1 = (idV[0]) ? (idV[0]) : ''
                let idV2 = (idV[1]) ? (idV[1]) : ''
                let idV3 = (idV[2]) ? (idV[2]) : ''
                let idV4 = (idV[3]) ? (idV[3]) : ''
                let idV5 = (idV[4]) ? (idV[4]) : ''
                let idV6 = (idV[5]) ? (idV[5]) : ''
                let idV7 = (idV[6]) ? (idV[6]) : ''
                let idV8 = (idV[7]) ? (idV[7]) : ''
                let idV9 = (idV[8]) ? (idV[8]) : ''
                let idV10 = (idV[9]) ? (idV[9]) : ''
                let idV11 = (idV[10]) ? (idV[10]) : ''
                let idV12 = (idV[11]) ? (idV[11]) : ''
                let idV13 = (idV[12]) ? (idV[12]) : ''
                let idV14 = (idV[13]) ? (idV[13]) : ''
                let idV15 = (idV[14]) ? (idV[14]) : ''

                // dateOfIssueVisa {dOIV1-8}
                let dOIV = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString().split(".") : ''
                let dOIV1 = (dOIV) ? (dOIV[0][0]) : ''
                let dOIV2 = (dOIV) ? (dOIV[0][1]) : ''
                let dOIV3 = (dOIV) ? (dOIV[1][0]) : ''
                let dOIV4 = (dOIV) ? (dOIV[1][1]) : ''
                let dOIV5 = (dOIV) ? (dOIV[2][0]) : ''
                let dOIV6 = (dOIV) ? (dOIV[2][1]) : ''
                let dOIV7 = (dOIV) ? (dOIV[2][2]) : ''
                let dOIV8 = (dOIV) ? (dOIV[2][3]) : ''

                // validUntilVisa {vUV1-8}
                let vUV = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString().split(".") : ''
                let vUV1 = (vUV) ? (vUV[0][0]) : ''
                let vUV2 = (vUV) ? (vUV[0][1]) : ''
                let vUV3 = (vUV) ? (vUV[1][0]) : ''
                let vUV4 = (vUV) ? (vUV[1][1]) : ''
                let vUV5 = (vUV) ? (vUV[2][0]) : ''
                let vUV6 = (vUV) ? (vUV[2][1]) : ''
                let vUV7 = (vUV) ? (vUV[2][2]) : ''
                let vUV8 = (vUV) ? (vUV[2][3]) : ''

                // purpose
                let purposeG = ""
                let purposeR = ""
                let purposeU = ""
                let purp1 = ''
                let purp2 = ''
                let purp3 = ''
                let purp4 = ''
                let purp5 = ''
                let purp6 = ''
                let purp7 = ''
                let purp8 = ''
                let purp9 = ''
                let purp10 = ''
                let purp11 = ''
                let purp12 = ''
                let purp13 = ''
                switch (document.getElementById('purpose' + indexTab).value) {
                    case "Учеба":
                        purposeU = "X"
                        purp1 = 'С'; purp2='Т'; purp3='У'; purp4="Д"; purp5 ='Е';purp6 = 'Н';purp7='Т'
                        break
                    case "Краткосрочная учеба":
                        purposeU = "X"
                        purp1 = 'С'; purp2='Т'; purp3='У'; purp4="Д"; purp5 ='Е';purp6 = 'Н';purp7='Т'
                        break
                    case "(НТС)":
                        purposeG = "X"
                        purp1 = 'Н'; purp2='Т'; purp3='С';
                        break
                    case "Трудовая деятельность":
                        purposeR = "X"
                        purp1='П';purp2='Р';purp3='Е';purp4="П";purp5='О';purp6='Д';purp7='А';purp8='В';purp9='А';purp10='Т';purp11='Е';purp12='Л';purp13='Ь';
                        break
                }

                // dateArrivalMigration {dAM1-8}
                let dAM = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString().split(".") : ''
                let dAM1 = (dAM) ? (dAM[0][0]) : ''
                let dAM2 = (dAM) ? (dAM[0][1]) : ''
                let dAM3 = (dAM) ? (dAM[1][0]) : ''
                let dAM4 = (dAM) ? (dAM[1][1]) : ''
                let dAM5 = (dAM) ? (dAM[2][0]) : ''
                let dAM6 = (dAM) ? (dAM[2][1]) : ''
                let dAM7 = (dAM) ? (dAM[2][2]) : ''
                let dAM8 = (dAM) ? (dAM[2][3]) : ''

                // dateUntil {dU1-8}
                let dU = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString().split(".") : ''

                let dU1 = (dU) ? (dU[0][0]) : ''
                let dU2 = (dU) ? (dU[0][1]) : ''
                let dU3 = (dU) ? (dU[1][0]) : ''
                let dU4 = (dU) ? (dU[1][1]) : ''
                let dU5 = (dU) ? (dU[2][0]) : ''
                let dU6 = (dU) ? (dU[2][1]) : ''
                let dU7 = (dU) ? (dU[2][2]) : ''
                let dU8 = (dU) ? (dU[2][3]) : ''

                // seriesMigration {sM1-4}
                let sM = document.getElementById('seriesMigration' + indexTab).value
                    ? document.getElementById('seriesMigration' + indexTab).value.toUpperCase() : ''
                let sM1 = ''
                let sM2 = ''
                let sM3 = ''
                let sM4 = ''
                if (sM) {
                    switch (sM.length) {
                        case 1:
                            sM4 = sM[0]
                            break
                        case 2:
                            sM3 = sM[0]
                            sM4 = sM[1]
                            break
                        case 3:
                            sM2 = sM[0]
                            sM3 = sM[1]
                            sM4 = sM[2]
                            break
                        case 4:
                            sM1 = sM[0]
                            sM2 = sM[1]
                            sM3 = sM[2]
                            sM4 = sM[3]
                    }
                }

                //idMigration {iM1-11}
                let iM = document.getElementById('idMigration' + indexTab).value
                    ? document.getElementById('idMigration' + indexTab).value.toUpperCase() : ''
                let iM1 = (iM[0]) ? (iM[0]) : ''
                let iM2 = (iM[1]) ? (iM[1]) : ''
                let iM3 = (iM[2]) ? (iM[2]) : ''
                let iM4 = (iM[3]) ? (iM[3]) : ''
                let iM5 = (iM[4]) ? (iM[4]) : ''
                let iM6 = (iM[5]) ? (iM[5]) : ''
                let iM7 = (iM[6]) ? (iM[6]) : ''
                let iM8 = (iM[7]) ? (iM[7]) : ''
                let iM9 = (iM[8]) ? (iM[8]) : ''
                let iM10 = (iM[9]) ? (iM[9]) : ''
                let iM11 = (iM[10]) ? (iM[10]) : ''

                // addressHostel {aHG1-18} {aHU1-20} {aHD1-15} {aHK1-5}
                let aHG1 = ''
                let aHG2 = ''
                let aHG3 = ''
                let aHG4 = ''
                let aHG5 = ''
                let aHG6 = ''
                let aHG7 = ''
                let aHG8 = ''
                let aHG9 = ''
                let aHG10 = ''
                let aHG11 = ''
                let aHG12 = ''
                let aHG13 = ''
                let aHG14 = ''
                let aHG15 = ''
                let aHG16 = ''
                let aHG17 = ''
                let aHG18 = ''
                let aHU1 = ''
                let aHU2 = ''
                let aHU3 = ''
                let aHU4 = ''
                let aHU5 = ''
                let aHU6 = ''
                let aHU7 = ''
                let aHU8 = ''
                let aHU9 = ''
                let aHU10 = ''
                let aHU11 = ''
                let aHU12 = ''
                let aHU13 = ''
                let aHU14 = ''
                let aHU15 = ''
                let aHU16 = ''
                let aHU17 = ''
                let aHU18 = ''
                let aHU19 = ''
                let aHU20 = ''
                let aHD1 = ""
                let aHD2 = ""
                let aHD3 = ""
                let aHD4 = ""
                let aHD5 = ""
                let aHD6 = ""
                let aHD7 = ""
                let aHD8 = ""
                let aHD9 = ""
                let aHD10 = ""
                let aHD11 = ""
                let aHD12 = ""
                let aHD13 = ""
                let aHD14 = ""
                let aHD15 = ""
                let aHK1 = ''
                let aHK2 = ''
                let aHK3 = ''
                let aHK4 = ''
                let aHK5 = ''

                switch (document.getElementById('addressHostel'+indexTab).value) {
                    case "г. Москва, проспект Вернадского, 88 к. 1 (ОБЩЕЖИТИЕ №1)":
                        aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                        aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                        aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '1';
                        break
                    case "г. Москва, проспект Вернадского, 88 к. 2 (ОБЩЕЖИТИЕ №2)":
                        aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                        aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                        aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '2';
                        break
                    case "г. Москва, проспект Вернадского, 88 к. 3 (ОБЩЕЖИТИЕ №3)":
                        aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                        aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                        aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '3';
                        break
                    case "г. Москва, улица Космонавтов, д. 13 (ОБЩЕЖИТИЕ №4)":
                        aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                        aHU1 = 'К'; aHU2 = 'О'; aHU3 = 'С'; aHU4 = 'М'; aHU5 = 'О'; aHU6 = 'Н'; aHU7 = 'А'; aHU8 = 'В'; aHU9 = 'Т'; aHU10 = 'О'; aHU11 = 'В';
                        aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '1'; aHD6 = '3';
                        break
                    case "г. Москва, улица Космонавтов, д. 9 (ОБЩЕЖИТИЕ №5)":
                        aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                        aHU1 = 'К'; aHU2 = 'О'; aHU3 = 'С'; aHU4 = 'М'; aHU5 = 'О'; aHU6 = 'Н'; aHU7 = 'А'; aHU8 = 'В'; aHU9 = 'Т'; aHU10 = 'О'; aHU11 = 'В';
                        aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '9';
                        break
                    case "г. Москва, улица Клары Цеткин, д. 25 (ОБЩЕЖИТИЕ №6)":
                        aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                        aHU1 = 'К'; aHU2 = 'Л'; aHU3 = 'А'; aHU4 = 'Р'; aHU5 = 'Ы'; aHU6 = ''; aHU7 = 'Ц'; aHU8 = 'Е'; aHU9 = 'Т'; aHU10 = 'К'; aHU11 = 'И'; aHU12 = 'Н';
                        aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '2'; aHD6 = '5';
                        break
                    case "Московская область, г. Люберцы, ул. Мира, д.7 (ОБЩЕЖИТИЕ №7)":
                        aHG1 = 'М'; aHG2 = 'О'; aHG3 = 'С'; aHG4 = 'К'; aHG5 = 'О'; aHG6 = 'В'; aHG7 = 'С'; aHG8 = 'К'; aHG9 = 'А'; aHG10 = 'Я'; aHG12 = 'О'; aHG13 = 'Б'; aHG14 = 'Л'; aHG15 = 'А'; aHG16 = 'С'; aHG17 = 'Т'; aHG18 = 'Ь';
                        aHU1 = 'Л'; aHU2 = 'Ю'; aHU3 = 'Б'; aHU4 = 'Е'; aHU5 = 'Р'; aHU6 = 'Ц'; aHU7 = 'Ы';
                        aHD1 = 'М'; aHD2 = 'И'; aHD3 = 'Р'; aHD4 = 'А';
                        aHK1 = 'Д'; aHK2 = 'О'; aHK3 = 'М'; aHK4 = ''; aHK5 = '7';
                        break
                }

                // migrationAddress {mAO1-18} {mAG1-7} {mAU1-20} {mAD1-2} {mAK1}
                let mAO1 = ""
                let mAO2 = ""
                let mAO3 = ""
                let mAO4 = ""
                let mAO5 = ""
                let mAO6 = ""
                let mAO7 = ""
                let mAO8 = ""
                let mAO9 = ""
                let mAO10 = ""
                let mAO11 = ""
                let mAO12 = ""
                let mAO13 = ""
                let mAO14 = ""
                let mAO15 = ""
                let mAO16 = ""
                let mAO17 = ""
                let mAO18 = ""
                let mAG1 = ''
                let mAG2 = ''
                let mAG3 = ''
                let mAG4 = ''
                let mAG5 = ''
                let mAG6 = ''
                let mAG7 = ''
                let mAU1 = ""
                let mAU2 = ""
                let mAU3 = ""
                let mAU4 = ""
                let mAU5 = ""
                let mAU6 = ""
                let mAU7 = ""
                let mAU8 = ""
                let mAU9 = ""
                let mAU10 = ""
                let mAU11 = ""
                let mAU12 = ""
                let mAU13 = ""
                let mAU14 = ""
                let mAU15 = ""
                let mAU16 = ""
                let mAU17 = ""
                let mAU18 = ""
                let mAU19 = ""
                let mAU20 = ""
                let mAD1 = ''
                let mAD2 = ''
                let mAK1 = ''
                switch (document.getElementById('migrationAddress').value) {
                    case 'г. Москва, проспект Вернадского, 88 к. 1':
                        mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                        mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                        mAD1 = '8'; mAD2 = '8'
                        mAK1 = '1'
                        break
                    case 'г. Москва, проспект Вернадского, 88 к. 2':
                        mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                        mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                        mAD1 = '8'; mAD2 = '8'
                        mAK1 = '2'
                        break
                    case 'г. Москва, проспект Вернадского, 88 к. 3':
                        mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                        mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                        mAD1 = '8'; mAD2 = '8'
                        mAK1 = '3'
                        break
                    case 'г. Москва, ул. Космонавтов, д. 13':
                        mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                        mAU1 = 'К'; mAU2 = 'О'; mAU3 = 'С'; mAU4 = 'М'; mAU5 = 'О'; mAU6 = 'Н'; mAU7 = 'А'; mAU8 = 'В'; mAU9 = 'Т'; mAU10 = 'О'; mAU11 = 'В';
                        mAD1 = '1'; mAD2 = '3'
                        break
                    case 'г. Москва, улица Космонавтов, д. 9':
                        mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                        mAU1 = 'К'; mAU2 = 'О'; mAU3 = 'С'; mAU4 = 'М'; mAU5 = 'О'; mAU6 = 'Н'; mAU7 = 'А'; mAU8 = 'В'; mAU9 = 'Т'; mAU10 = 'О'; mAU11 = 'В';
                        mAD2 = '9'
                        break
                    case 'г. Москва, ул. Клары Цеткин, д. 25, корп. 1':
                        mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                        mAU1 = 'К'; mAU2 = 'Л'; mAU3 = 'А'; mAU4 = 'Р'; mAU5 = 'Ы'; mAU7 = 'Ц'; mAU8 = 'Е'; mAU9 = 'Т'; mAU10 = 'К'; mAU11 = 'И'; mAU12 = 'Н';
                        mAD1 = '2'; mAD2 = '5'
                        mAK1 = '1'
                        break
                    case 'Московская область, г. Люберцы, ул. Мира, д.7':
                        mAO1 = 'М'; mAO2 = 'О'; mAO3 = 'С'; mAO4 = 'К'; mAO5 = 'О'; mAO6 = 'В'; mAO7 = 'С'; mAO8 = 'К'; mAO9 = 'А'; mAO10 = 'Я'; mAO12 = 'О'; mAO13 = 'Б'; mAO14 = 'Л'; mAO15 = 'А'; mAO16 = 'С'; mAO17 = 'Т'; mAO18 = 'Ь';
                        mAG1 = 'Л'; mAG2 = 'Ю'; mAG3 = 'Б'; mAG4 = 'Е'; mAG5 = 'Р'; mAG6 = 'Ц'; mAG7 = 'Ы';
                        mAD2 = '7'
                        break
                }


                //numRoom {nR1-4}
                let nR = document.getElementById('numRoom'+ indexTab).value != "-"
                    ? document.getElementById('numRoom' + indexTab).value.toUpperCase() : ''
                let nR1 = ''
                let nR2 = ''
                let nR3 = ''
                let nR4 = ''
                if (nR) {
                    switch (nR.length) {
                        case 1:
                            nR4 = nR[0]
                            break
                        case 2:
                            nR3 = nR[0]
                            nR4 = nR[1]
                            break
                        case 3:
                            nR2 = nR[0]
                            nR3 = nR[1]
                            nR4 = nR[2]
                            break
                        case 4:
                            nR1 = nR[0]
                            nR2 = nR[1]
                            nR3 = nR[2]
                            nR4 = nR[3]
                    }
                }






                doc.setData({
                    lN1: lN1, lN2: lN2, lN3: lN3, lN4: lN4, lN5: lN5, lN6: lN6, lN7: lN7, lN8: lN8, lN9: lN9, lN10: lN10, lN11: lN11, lN12: lN12, lN13: lN13, lN14: lN14, lN15: lN15, lN16: lN16, lN17: lN17, lN18: lN18, lN19: lN19, lN20: lN20, lN21: lN21, lN22: lN22, lN23: lN23, lN24: lN24, lN25: lN25, lN26: lN26, lN27: lN27,
                    fN1: fN1, fN2: fN2, fN3: fN3, fN4: fN4, fN5: fN5, fN6: fN6, fN7: fN7, fN8: fN8, fN9: fN9, fN10: fN10, fN11: fN11, fN12: fN12, fN13: fN13, fN14: fN14, fN15: fN15, fN16: fN16, fN17: fN17, fN18: fN18, fN19: fN19, fN20: fN20, fN21: fN21, fN22: fN22, fN23: fN23, fN24: fN24, fN25: fN25, fN26: fN26, fN27: fN27,
                    patr1: patr1, patr2: patr2, patr3: patr3, patr4: patr4, patr5: patr5, patr6: patr6, patr7: patr7, patr8: patr8, patr9: patr9, patr10: patr10, patr11: patr11, patr12: patr12, patr13: patr13, patr14: patr14, patr15: patr15, patr16: patr16, patr17: patr17, patr18: patr18, patr19: patr19, patr20: patr20, patr21: patr21, patr22: patr22, patr23: patr23, patr24: patr24,
                    grazd1: grazd1, grazd2: grazd2, grazd3: grazd3, grazd4: grazd4, grazd5: grazd5, grazd6: grazd6, grazd7: grazd7, grazd8: grazd8, grazd9: grazd9, grazd10: grazd10, grazd11: grazd11, grazd12: grazd12, grazd13: grazd13, grazd14: grazd14, grazd15: grazd15, grazd16: grazd16, grazd17: grazd17, grazd18: grazd18, grazd19: grazd19, grazd20: grazd20, grazd21: grazd21, grazd22: grazd22, grazd23: grazd23, grazd24: grazd24, grazd25: grazd25,
                    dOB1: dOB1, dOB2: dOB2, dOB3: dOB3, dOB4: dOB4, dOB5: dOB5, dOB6: dOB6, dOB7: dOB7, dOB8: dOB8,
                    genM: genM, genW: genW,
                    pSBS1: pSBS1, pSBS2: pSBS2, pSBS3: pSBS3, pSBS4: pSBS4, pSBS5: pSBS5, pSBS6: pSBS6, pSBS7: pSBS7, pSBS8: pSBS8, pSBS9: pSBS9, pSBS10: pSBS10, pSBS11: pSBS11, pSBS12: pSBS12, pSBS13: pSBS13, pSBS14: pSBS14, pSBS15: pSBS15, pSBS16: pSBS16, pSBS17: pSBS17, pSBS18: pSBS18, pSBS19: pSBS19, pSBS20: pSBS20, pSBS21: pSBS21, pSBS22: pSBS22, pSBS23: pSBS23, pSBS24: pSBS24,
                    pSBG1: pSBG1, pSBG2: pSBG2, pSBG3: pSBG3, pSBG4: pSBG4, pSBG5: pSBG5, pSBG6: pSBG6, pSBG7: pSBG7, pSBG8: pSBG8, pSBG9: pSBG9, pSBG10: pSBG10, pSBG11: pSBG11, pSBG12: pSBG12, pSBG13: pSBG13, pSBG14: pSBG14, pSBG15: pSBG15, pSBG16: pSBG16, pSBG17: pSBG17, pSBG18: pSBG18, pSBG19: pSBG19, pSBG20: pSBG20, pSBG21: pSBG21, pSBG22: pSBG22, pSBG23: pSBG23, pSBG24: pSBG24,
                    sP1: sP1, sP2: sP2, sP3: sP3, sP4: sP4,
                    idP1: idP1, idP2: idP2, idP3: idP3, idP4: idP4, idP5: idP5, idP6: idP6, idP7: idP7, idP8: idP8, idP9: idP9, idP10: idP10,
                    dOI1: dOI1, dOI2: dOI2, dOI3: dOI3, dOI4: dOI4, dOI5: dOI5, dOI6: dOI6, dOI7: dOI7, dOI8: dOI8,
                    vU1: vU1, vU2: vU2, vU3: vU3, vU4: vU4, vU5: vU5, vU6: vU6, vU7: vU7, vU8: vU8,
                    tVV: tVV, tVJ: tVJ, tVP: tVP,
                    sV1: sV1, sV2: sV2, sV3: sV3, sV4: sV4,
                    idV1: idV1, idV2: idV2, idV3: idV3, idV4: idV4, idV5: idV5, idV6: idV6, idV7: idV7, idV8: idV8, idV9: idV9, idV10: idV10, idV11: idV11, idV12: idV12, idV13: idV13, idV14: idV14, idV15: idV15,
                    dOIV1: dOIV1, dOIV2: dOIV2, dOIV3: dOIV3, dOIV4: dOIV4, dOIV5: dOIV5, dOIV6: dOIV6, dOIV7: dOIV7, dOIV8: dOIV8,
                    vUV1: vUV1, vUV2: vUV2, vUV3: vUV3, vUV4: vUV4, vUV5: vUV5, vUV6: vUV6, vUV7: vUV7, vUV8: vUV8,
                    dAM1: dAM1, dAM2: dAM2, dAM3: dAM3, dAM4: dAM4, dAM5: dAM5, dAM6: dAM6, dAM7: dAM7, dAM8: dAM8,
                    dU1: dU1, dU2: dU2, dU3: dU3, dU4: dU4, dU5: dU5, dU6: dU6, dU7: dU7, dU8: dU8,
                    sM1: sM1, sM2: sM2, sM3: sM3, sM4: sM4,
                    iM1: iM1, iM2: iM2, iM3: iM3, iM4: iM4, iM5: iM5, iM6: iM6, iM7: iM7, iM8: iM8, iM9: iM9, iM10: iM10, iM11: iM11,
                    aHG1: aHG1, aHG2: aHG2, aHG3: aHG3, aHG4: aHG4, aHG5: aHG5, aHG6: aHG6, aHG7: aHG7, aHG8: aHG8, aHG9: aHG9, aHG10: aHG10, aHG11: aHG11, aHG12: aHG12, aHG13: aHG13, aHG14: aHG14, aHG15: aHG15, aHG16: aHG16, aHG17: aHG17, aHG18: aHG18,
                    aHU1: aHU1, aHU2: aHU2, aHU3: aHU3, aHU4: aHU4, aHU5: aHU5, aHU6: aHU6, aHU7: aHU7, aHU8: aHU8, aHU9: aHU9, aHU10: aHU10, aHU11: aHU11, aHU12: aHU12, aHU13: aHU13, aHU14: aHU14, aHU15: aHU15, aHU16: aHU16, aHU17: aHU17, aHU18: aHU18, aHU19: aHU19, aHU20: aHU20,
                    aHD1: aHD1, aHD2: aHD2, aHD3: aHD3, aHD4: aHD4, aHD5: aHD5, aHD6: aHD6, aHD7: aHD7, aHD8: aHD8, aHD9: aHD9, aHD10: aHD10, aHD11: aHD11, aHD12: aHD12, aHD13: aHD13, aHD14: aHD14, aHD15: aHD15,
                    aHK1: aHK1, aHK2: aHK2, aHK3: aHK3, aHK4: aHK4, aHK5: aHK5,
                    mAO1: mAO1, mAO2: mAO2, mAO3: mAO3, mAO4: mAO4, mAO5: mAO5, mAO6: mAO6, mAO7: mAO7, mAO8: mAO8, mAO9: mAO9, mAO10: mAO10, mAO11: mAO11, mAO12: mAO12, mAO13: mAO13, mAO14: mAO14, mAO15: mAO15, mAO16: mAO16, mAO17: mAO17, mAO18: mAO18,
                    mAG1: mAG1, mAG2: mAG2, mAG3: mAG3, mAG4: mAG4, mAG5: mAG5, mAG6: mAG6, mAG7: mAG7,
                    mAU1: mAU1, mAU2: mAU2, mAU3: mAU3, mAU4: mAU4, mAU5: mAU5, mAU6: mAU6, mAU7: mAU7, mAU8: mAU8, mAU9: mAU9, mAU10: mAU10, mAU11: mAU11, mAU12: mAU12, mAU13: mAU13, mAU14: mAU14, mAU15: mAU15, mAU16: mAU16, mAU17: mAU17, mAU18: mAU18, mAU19: mAU19, mAU20: mAU20,
                    mAD1: mAD1, mAD2: mAD2, mAK1: mAK1,


                    purp1: purp1, purp2: purp2, purp3: purp3, purp4: purp4, purp5: purp5, purp6: purp6, purp7: purp7, purp8: purp8, purp9: purp9, purp10: purp10, purp11: purp11, purp12: purp12, purp13: purp13,
                    purposeG: purposeG, purposeU: purposeU, purposeR: purposeR,

                    nR1: nR1, nR2: nR2, nR3: nR3, nR4: nR4,


                });



                try {
                    doc.render();
                }
                catch (error) {
                    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                    errorHandler(error);
                }
                var out = doc.getZip().generate();
                zipDocs.file("УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                    document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                    " - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
                    , out, {base64: true}
                );
            } // end for


            let nameFile = ''
            if (countTab()==1) {
                nameFile = document.getElementById('nStud1').value
                    +" УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    +".zip"
            }
            else {
                nameFile = document.getElementById('nStud1').value + '-'+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".zip"
            }
            var content = zipDocs.generate({ type: "blob" });
            saveAs(content,nameFile);
        });
};






//регистрация+виза

//опись
window.generateInventoryRegVisa = function generate() {
    path = ('../Templates/регистрация И виза/опись.docx')

    let students = []

    // registration On
    let registrationOn = ''
    switch (document.getElementById('registrationOn').value) {
        case "Круглов":
            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
            break
        case "Морозова":
            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
            break
        case "Орлова":
            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
            break
    }


    // nStud
    let nStud1 = document.getElementById('nStud1').value
    let tir = ''
    let nStud2 = ''
    if (countTab()>1) {
        nStud2 = document.getElementById('nStud' + (lastTab()-1)).value
        tir = '-'
    }

    for (let i =0; i<countTab();i++) {
        let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
        let elem = tabs[i]
        let indexTab = parseInt(elem.id.match(/\d+/))


        let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
            ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''

        let dateUntil = document.getElementById('dateUntil' + indexTab).value != '-' ?
            new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

        //students
        students.push({
            nStud: document.getElementById('nStud' + indexTab).value,
            lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
            firstNameRu: document.getElementById('firstNameRu'+indexTab).value.toUpperCase(),
            patronymicRu: document.getElementById('patronymicRu'+indexTab).value.toUpperCase(),
            dateOfIssueVisa: dateOfIssueVisa,
            dateUntil: dateUntil,
            grazd: document.getElementById('grazd' + indexTab).value,
            phone: document.getElementById('phone' + indexTab).value,
            mail: document.getElementById('mail' + indexTab).value,
        })
    }


    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }
            var zip = new PizZip(content);
            var doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            doc.render({
                dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                nStud1: nStud1,
                nStud2: nStud2,
                tir: tir,
                'students': students,
                registrationOn: registrationOn,
            })


            var out = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                // compression: DEFLATE adds a compression step.
                // For a 50MB output document, expect 500ms additional CPU time
                compression: "DEFLATE",
            });



            let nameFile = ''
            if (countTab()==1) {
                nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ + ВИЗА) - студент " +
                    document.getElementById('nStud1').value
                    +" - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                    +" - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
            }
            else {
                nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ + ВИЗА) - студенты " +
                    document.getElementById('nStud1').value
                    +"-"+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                    +" - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
            }


            // Output the document using Data-URI
            saveAs(out, nameFile);
        }
    );
};





//ходатайство по квартире
window.generateFlatSolicitaion = function generate() {
    path = ('../Templates/ходатайство по квартире.docx')

    var zipDocs = new PizZip();
    loadFile(
        path,
        function (error, content) {
            if (error) {
                throw error;
            }

            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }
            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            for (let i =0; i<countTab();i++) {
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                let elem = tabs[i]
                let indexTab = parseInt(elem.id.match(/\d+/))


                //dateUntil
                let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                // purpose
                let purpose = ''
                let purposeS = ''
                switch (document.getElementById('purpose' + indexTab).value) {
                    case "Учеба":
                        purpose = "УЧЕБА"
                        purposeS = "Студент"
                        break
                    case "Краткосрочная учеба":
                        purpose = "КРАТКОСРОЧНАЯ УЧЕБА"
                        purposeS = "Студент"
                        break
                    case "(НТС)":
                        purpose = "НАУЧНО-ТЕХНИЧЕСКИЕ СВЯЗИ (НТС)"
                        purposeS = "НТС"
                        break
                    case "Трудовая деятельность":
                        purpose = "ТРУДОВАЯ ДЕЯТЕЛЬНОСТЬ"
                        purposeS = "Преподаватель"
                        break
                }

                // levelEducation
                let levelEducation = ''
                switch (document.getElementById('levelEducation' + indexTab).value) {
                    case "Подготовительный факультет (изучаю русский язык)/ The preparatory faculty":
                        levelEducation = 'подготовительный факультет'
                        break
                    case "бакалавриат/bachelor degree":
                        levelEducation = 'бакалавриат'
                        break
                    case "магистратура/master degree":
                        levelEducation = 'магистратура'
                        break
                    case "аспирантура/post-graduate studies":
                        levelEducation = 'аспирантура'
                        break
                }

                // course
                let course = ''
                switch (document.getElementById('course' + indexTab).value) {
                    case '1':
                        course = ', 1 курс,'
                        break
                    case '2':
                        course = ', 2 курс,'
                        break
                    case '3':
                        course = ', 3 курс,'
                        break
                    case '4':
                        course = ', 4 курс,'
                        break
                    case '5':
                        course = ', 5 курс,'
                        break
                }

                // addressResidence
                let addressResidence = ''
                switch (document.getElementById('migrationAddress').value) {
                    case "Квартира":
                        addressResidence = document.getElementById('addressResidence' + indexTab).value
                        break
                    default:
                        addressResidence = document.getElementById('migrationAddress').value
                        break
                }

                // Passport
                let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                    ? document.getElementById('series' + indexTab).value : ''
                let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                // gender
                let gender = ''
                switch (document.getElementById('gender' + indexTab).value) {
                    case "Мужской / Male":
                        gender = 'м.'
                        break
                    case "Женский / Female":
                        gender = 'ж.'
                        break
                }

                // registration On
                let registrationOn = ''
                switch (document.getElementById('registrationOn').value) {
                    case "Круглов":
                        registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                        break
                    case "Морозова":
                        registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                        break
                    case "Орлова":
                        registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                        break
                }

                // visa
                let typeVisa = ''
                switch (document.getElementById('typeVisa' + indexTab).value) {
                    case "ВИЗА":
                        typeVisa = 'Виза'
                        break
                    case "(ВНЖ) ВИД НА ЖИТЕЛЬСТВО РФ":
                        typeVisa = 'ВНЖ'
                        break
                    case "(РВП) РАЗРЕШЕНИЕ НА ВРЕМЕННОЕ ПРОЖИВАНИЕ РФ":
                        typeVisa = 'РВП'
                        break

                }
                let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                    ? document.getElementById('seriesVisa' + indexTab).value : ''
                let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                    ? document.getElementById('idVisa' + indexTab).value : ''

                let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                    ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''


                // faculty
                let faculty = ''
                switch (document.getElementById('faculty' + indexTab).value) {
                    case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                        faculty = 'ИИИ:Музфак'
                        break
                    case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                        faculty = 'ИИИ: Худграф'
                        break
                    case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                        faculty = 'ИСГО'
                        break
                    case "Институт филологии / The Institute of Philology":
                        faculty = 'ИФ'
                        break
                    case "Институт иностранных языков / The Institute of Foreign Languages":
                        faculty = 'ИИЯ'
                        break
                    case "Институт международного образования / The Institute of International Education":
                        faculty = 'ИМО'
                        break
                    case "Институт детства / The Institute of Childhood":
                        faculty = 'ИД'
                        break
                    case "Институт биологии и химии / The Institute of Biology and Chemistry":
                        faculty = 'ИБХ'
                        break
                    case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                        faculty = 'ИФТИС'
                        break
                    case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                        faculty = 'ИФКСиЗ'
                        break
                    case "Географический факультет / The Institute of Geography":
                        faculty = 'Геофак'
                        break
                    case "Институт истории и политики / The Institute of History and Politics":
                        faculty = 'ИИП'
                        break
                    case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                        faculty = 'ИМИ'
                        break
                    case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                        faculty = 'Дош.фак.'
                        break
                    case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                        faculty = 'ИПП'
                        break
                    case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                        faculty = 'ИЖКиМ'
                        break
                    case "Институт развития цифрового образования / The Institute of Digital Education Development":
                        faculty = 'ИРЦО'
                        break
                }


                doc.setData({
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    nStud: document.getElementById('nStud' + indexTab).value,
                    grazd: document.getElementById('grazd' + indexTab).value,
                    lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                    firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                    patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                    dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                    gender: gender,

                    registrationOn: registrationOn,
                    dateUntil: dateUntil,
                    purpose: purpose,
                    purposeS: purposeS,
                    levelEducation: levelEducation,
                    course: course,

                    series: series,
                    idPassport: document.getElementById('idPassport' + indexTab).value,
                    dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                    validUntil: validUntil,

                    typeVisa: typeVisa,
                    seriesVisa: seriesVisa,
                    idVisa: idVisa,
                    dateOfIssueVisa: dateOfIssueVisa,
                    validUntilVisa: validUntilVisa,

                    seriesMigration: document.getElementById('seriesMigration' + indexTab).value,
                    idMigration: document.getElementById('idMigration' + indexTab).value,
                    dateArrivalMigration: new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString(),

                    addressResidence: document.getElementById('addressResidence' + indexTab).value,
                });



                try {
                    doc.render();
                }
                catch (error) {
                    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                    errorHandler(error);
                }
                var out = doc.getZip().generate();
                zipDocs.file("ХОДАТАЙСТВО ПО КВАРТИРЕ - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                    document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                    document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                    new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                    " - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".docx"
                    , out, {base64: true}
                );
            } // end for


            let nameFile = ''
            if (countTab()==1) {
                nameFile = document.getElementById('nStud1').value
                    +" ХОДАТАЙСТВО ПО КВАРТИРЕ - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    +".zip"
            }
            else {
                nameFile = document.getElementById('nStud1').value + '-'+
                    document.getElementById('nStud'+(lastTab()-1)).value
                    +" ХОДАТАЙСТВО ПО КВАРТИРЕ - " +
                    document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                    + ".zip"
            }
            var content = zipDocs.generate({ type: "blob" });
            saveAs(content,nameFile);
        });
};














// РЕГИСТРАЦИЯ общий выгруз
function generateReg() {
    var zipTotal = new PizZip();
//уведомление РЕГИСТРАЦИЯ
    window.generateRegNotifTotal = function generate() {
        let ovmRg = document.getElementById('ovmByRegion').value
        let rgOn = document.getElementById('registrationOn').value
        path = (`../Templates/регистрация/уведомление ${ovmRg} ${rgOn}.docx`)

        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))

                    // lastName {lN1-27}
                    let lN = document.getElementById('lastNameRu'+indexTab).value.toUpperCase().split('')
                    let lN1 = (lN[0]) ? (lN[0]) : ''
                    let lN2 = (lN[1]) ? (lN[1]) : ''
                    let lN3 = (lN[2]) ? (lN[2]) : ''
                    let lN4 = (lN[3]) ? (lN[3]) : ''
                    let lN5 = (lN[4]) ? (lN[4]) : ''
                    let lN6 = (lN[5]) ? (lN[5]) : ''
                    let lN7 = (lN[6]) ? (lN[6]) : ''
                    let lN8 = (lN[7]) ? (lN[7]) : ''
                    let lN9 = (lN[8]) ? (lN[8]) : ''
                    let lN10 = (lN[9]) ? (lN[9]) : ''
                    let lN11 = (lN[10]) ? (lN[10]) : ''
                    let lN12 = (lN[11]) ? (lN[11]) : ''
                    let lN13 = (lN[12]) ? (lN[12]) : ''
                    let lN14 = (lN[13]) ? (lN[13]) : ''
                    let lN15 = (lN[14]) ? (lN[14]) : ''
                    let lN16 = (lN[15]) ? (lN[15]) : ''
                    let lN17 = (lN[16]) ? (lN[16]) : ''
                    let lN18 = (lN[17]) ? (lN[17]) : ''
                    let lN19 = (lN[18]) ? (lN[18]) : ''
                    let lN20 = (lN[19]) ? (lN[19]) : ''
                    let lN21 = (lN[20]) ? (lN[20]) : ''
                    let lN22 = (lN[21]) ? (lN[21]) : ''
                    let lN23 = (lN[22]) ? (lN[22]) : ''
                    let lN24 = (lN[23]) ? (lN[23]) : ''
                    let lN25 = (lN[24]) ? (lN[24]) : ''
                    let lN26 = (lN[25]) ? (lN[25]) : ''
                    let lN27 = (lN[26]) ? (lN[26]) : ''

                    // firstName {fN1-f27}
                    let fN = document.getElementById('firstNameRu' + indexTab).value.toUpperCase().split('')
                    let fN1 = (fN[0]) ? (fN[0]) : ''
                    let fN2 = (fN[1]) ? (fN[1]) : ''
                    let fN3 = (fN[2]) ? (fN[2]) : ''
                    let fN4 = (fN[3]) ? (fN[3]) : ''
                    let fN5 = (fN[4]) ? (fN[4]) : ''
                    let fN6 = (fN[5]) ? (fN[5]) : ''
                    let fN7 = (fN[6]) ? (fN[6]) : ''
                    let fN8 = (fN[7]) ? (fN[7]) : ''
                    let fN9 = (fN[8]) ? (fN[8]) : ''
                    let fN10 = (fN[9]) ? (fN[9]) : ''
                    let fN11 = (fN[10]) ? (fN[10]) : ''
                    let fN12 = (fN[11]) ? (fN[11]) : ''
                    let fN13 = (fN[12]) ? (fN[12]) : ''
                    let fN14 = (fN[13]) ? (fN[13]) : ''
                    let fN15 = (fN[14]) ? (fN[14]) : ''
                    let fN16 = (fN[15]) ? (fN[15]) : ''
                    let fN17 = (fN[16]) ? (fN[16]) : ''
                    let fN18 = (fN[17]) ? (fN[17]) : ''
                    let fN19 = (fN[18]) ? (fN[18]) : ''
                    let fN20 = (fN[19]) ? (fN[19]) : ''
                    let fN21 = (fN[20]) ? (fN[20]) : ''
                    let fN22 = (fN[21]) ? (fN[21]) : ''
                    let fN23 = (fN[22]) ? (fN[22]) : ''
                    let fN24 = (fN[23]) ? (fN[23]) : ''
                    let fN25 = (fN[24]) ? (fN[24]) : ''
                    let fN26 = (fN[25]) ? (fN[25]) : ''
                    let fN27 = (fN[26]) ? (fN[26]) : ''

                    // patronymic {patr1-24}
                    let patr = document.getElementById('patronymicRu' + indexTab).value.toUpperCase().split('')
                    let patr1 = (patr[0]) ? (patr[0]) : ''
                    let patr2 = (patr[1]) ? (patr[1]) : ''
                    let patr3 = (patr[2]) ? (patr[2]) : ''
                    let patr4 = (patr[3]) ? (patr[3]) : ''
                    let patr5 = (patr[4]) ? (patr[4]) : ''
                    let patr6 = (patr[5]) ? (patr[5]) : ''
                    let patr7 = (patr[6]) ? (patr[6]) : ''
                    let patr8 = (patr[7]) ? (patr[7]) : ''
                    let patr9 = (patr[8]) ? (patr[8]) : ''
                    let patr10 = (patr[9]) ? (patr[9]) : ''
                    let patr11 = (patr[10]) ? (patr[10]) : ''
                    let patr12 = (patr[11]) ? (patr[11]) : ''
                    let patr13 = (patr[12]) ? (patr[12]) : ''
                    let patr14 = (patr[13]) ? (patr[13]) : ''
                    let patr15 = (patr[14]) ? (patr[14]) : ''
                    let patr16 = (patr[15]) ? (patr[15]) : ''
                    let patr17 = (patr[16]) ? (patr[16]) : ''
                    let patr18 = (patr[17]) ? (patr[17]) : ''
                    let patr19 = (patr[18]) ? (patr[18]) : ''
                    let patr20 = (patr[19]) ? (patr[19]) : ''
                    let patr21 = (patr[20]) ? (patr[20]) : ''
                    let patr22 = (patr[21]) ? (patr[21]) : ''
                    let patr23 = (patr[22]) ? (patr[22]) : ''
                    let patr24 = (patr[23]) ? (patr[23]) : ''

                    // grazd {grazd1-25}
                    let grazd = document.getElementById('grazd' + indexTab).value.toUpperCase().split('')
                    let grazd1 = (grazd[0]) ? (grazd[0]) : ''
                    let grazd2 = (grazd[1]) ? (grazd[1]) : ''
                    let grazd3 = (grazd[2]) ? (grazd[2]) : ''
                    let grazd4 = (grazd[3]) ? (grazd[3]) : ''
                    let grazd5 = (grazd[4]) ? (grazd[4]) : ''
                    let grazd6 = (grazd[5]) ? (grazd[5]) : ''
                    let grazd7 = (grazd[6]) ? (grazd[6]) : ''
                    let grazd8 = (grazd[7]) ? (grazd[7]) : ''
                    let grazd9 = (grazd[8]) ? (grazd[8]) : ''
                    let grazd10 = (grazd[9]) ? (grazd[9]) : ''
                    let grazd11 = (grazd[10]) ? (grazd[10]) : ''
                    let grazd12 = (grazd[11]) ? (grazd[11]) : ''
                    let grazd13 = (grazd[12]) ? (grazd[12]) : ''
                    let grazd14 = (grazd[13]) ? (grazd[13]) : ''
                    let grazd15 = (grazd[14]) ? (grazd[14]) : ''
                    let grazd16 = (grazd[15]) ? (grazd[15]) : ''
                    let grazd17 = (grazd[16]) ? (grazd[16]) : ''
                    let grazd18 = (grazd[17]) ? (grazd[17]) : ''
                    let grazd19 = (grazd[18]) ? (grazd[18]) : ''
                    let grazd20 = (grazd[19]) ? (grazd[19]) : ''
                    let grazd21 = (grazd[20]) ? (grazd[20]) : ''
                    let grazd22 = (grazd[21]) ? (grazd[21]) : ''
                    let grazd23 = (grazd[22]) ? (grazd[22]) : ''
                    let grazd24 = (grazd[23]) ? (grazd[23]) : ''
                    let grazd25 = (grazd[24]) ? (grazd[24]) : ''

                    // dateOfBirth {dOB1-8}
                    let dOB = new Date(document.getElementById('dateOfBirth'+indexTab).value).toLocaleDateString().split('.')
                    let dOB1 = (dOB) ? (dOB[0][0]) : ''
                    let dOB2 = (dOB) ? (dOB[0][1]) : ''
                    let dOB3 = (dOB) ? (dOB[1][0]) : ''
                    let dOB4 = (dOB) ? (dOB[1][1]) : ''
                    let dOB5 = (dOB) ? (dOB[2][0]) : ''
                    let dOB6 = (dOB) ? (dOB[2][1]) : ''
                    let dOB7 = (dOB) ? (dOB[2][2]) : ''
                    let dOB8 = (dOB) ? (dOB[2][3]) : ''

                    // gender
                    let genM = ''
                    let genW = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            genM = 'X'
                            break
                        case "Женский / Female":
                            genW = 'X'
                            break
                    }

                    // placeStateBirth {pSBS1-24} для страны И {pSBG1-24} для города
                    let pSBS = document.getElementById('placeStateBirth'+indexTab).value.toUpperCase().split(' ')
                    let pSBS1 = (pSBS[0][0]) ? (pSBS[0][0]) : ''
                    let pSBS2 = (pSBS[0][1]) ? (pSBS[0][1]) : ''
                    let pSBS3 = (pSBS[0][2]) ? (pSBS[0][2]) : ''
                    let pSBS4 = (pSBS[0][3]) ? (pSBS[0][3]) : ''
                    let pSBS5 = (pSBS[0][4]) ? (pSBS[0][4]) : ''
                    let pSBS6 = (pSBS[0][5]) ? (pSBS[0][5]) : ''
                    let pSBS7 = (pSBS[0][6]) ? (pSBS[0][6]) : ''
                    let pSBS8 = (pSBS[0][7]) ? (pSBS[0][7]) : ''
                    let pSBS9 = (pSBS[0][8]) ? (pSBS[0][8]) : ''
                    let pSBS10 = (pSBS[0][9]) ? (pSBS[0][9]) : ''
                    let pSBS11 = (pSBS[0][10]) ? (pSBS[0][10]) : ''
                    let pSBS12 = (pSBS[0][11]) ? (pSBS[0][11]) : ''
                    let pSBS13 = (pSBS[0][12]) ? (pSBS[0][12]) : ''
                    let pSBS14 = (pSBS[0][13]) ? (pSBS[0][13]) : ''
                    let pSBS15 = (pSBS[0][14]) ? (pSBS[0][14]) : ''
                    let pSBS16 = (pSBS[0][15]) ? (pSBS[0][15]) : ''
                    let pSBS17 = (pSBS[0][16]) ? (pSBS[0][16]) : ''
                    let pSBS18 = (pSBS[0][17]) ? (pSBS[0][17]) : ''
                    let pSBS19 = (pSBS[0][18]) ? (pSBS[0][18]) : ''
                    let pSBS20 = (pSBS[0][19]) ? (pSBS[0][19]) : ''
                    let pSBS21 = (pSBS[0][20]) ? (pSBS[0][20]) : ''
                    let pSBS22 = (pSBS[0][21]) ? (pSBS[0][21]) : ''
                    let pSBS23 = (pSBS[0][22]) ? (pSBS[0][22]) : ''
                    let pSBS24 = (pSBS[0][23]) ? (pSBS[0][23]) : ''
                    let pSBG1 = ""
                    let pSBG2 = ""
                    let pSBG3 = ""
                    let pSBG4 = ""
                    let pSBG5 = ""
                    let pSBG6 = ""
                    let pSBG7 = ""
                    let pSBG8 = ""
                    let pSBG9 = ""
                    let pSBG10 = ""
                    let pSBG11 = ""
                    let pSBG12 = ""
                    let pSBG13 = ""
                    let pSBG14 = ""
                    let pSBG15 = ""
                    let pSBG16 = ""
                    let pSBG17 = ""
                    let pSBG18 = ""
                    let pSBG19 = ""
                    let pSBG20 = ""
                    let pSBG21 = ""
                    let pSBG22 = ""
                    let pSBG23 = ""
                    let pSBG24 = ""
                    if (pSBS.length>1) {
                        pSBG1 = (pSBS[1][0]) ? (pSBS[1][0]) : ''
                        pSBG2 = (pSBS[1][1]) ? (pSBS[1][1]) : ''
                        pSBG3 = (pSBS[1][2]) ? (pSBS[1][2]) : ''
                        pSBG4 = (pSBS[1][3]) ? (pSBS[1][3]) : ''
                        pSBG5 = (pSBS[1][4]) ? (pSBS[1][4]) : ''
                        pSBG6 = (pSBS[1][5]) ? (pSBS[1][5]) : ''
                        pSBG7 = (pSBS[1][6]) ? (pSBS[1][6]) : ''
                        pSBG8 = (pSBS[1][7]) ? (pSBS[1][7]) : ''
                        pSBG9 = (pSBS[1][8]) ? (pSBS[1][8]) : ''
                        pSBG10 = (pSBS[1][9]) ? (pSBS[1][9]) : ''
                        pSBG11 = (pSBS[1][10]) ? (pSBS[1][10]) : ''
                        pSBG12 = (pSBS[1][11]) ? (pSBS[1][11]) : ''
                        pSBG13 = (pSBS[1][12]) ? (pSBS[1][12]) : ''
                        pSBG14 = (pSBS[1][13]) ? (pSBS[1][13]) : ''
                        pSBG15 = (pSBS[1][14]) ? (pSBS[1][14]) : ''
                        pSBG16 = (pSBS[1][15]) ? (pSBS[1][15]) : ''
                        pSBG17 = (pSBS[1][16]) ? (pSBS[1][16]) : ''
                        pSBG18 = (pSBS[1][17]) ? (pSBS[1][17]) : ''
                        pSBG19 = (pSBS[1][18]) ? (pSBS[1][18]) : ''
                        pSBG20 = (pSBS[1][19]) ? (pSBS[1][19]) : ''
                        pSBG21 = (pSBS[1][20]) ? (pSBS[1][20]) : ''
                        pSBG22 = (pSBS[1][21]) ? (pSBS[1][21]) : ''
                        pSBG23 = (pSBS[1][22]) ? (pSBS[1][22]) : ''
                        pSBG24 = (pSBS[1][23]) ? (pSBS[1][23]) : ''
                    }

                    // series {sP1-4}
                    let sP = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value.toUpperCase() : ''
                    let sP1 = ''
                    let sP2 = ''
                    let sP3 = ''
                    let sP4 = ''
                    if (sP) {
                        switch (sP.length) {
                            case 1:
                                sP4 = sP[0]
                                break
                            case 2:
                                sP3 = sP[0]
                                sP4 = sP[1]
                                break
                            case 3:
                                sP2 = sP[0]
                                sP3 = sP[1]
                                sP4 = sP[2]
                                break
                            case 4:
                                sP1 = sP[0]
                                sP2 = sP[1]
                                sP3 = sP[2]
                                sP4 = sP[3]
                        }
                    }


                    // idPassport {iP1-10}
                    let idP = document.getElementById('idPassport'+indexTab).value.toUpperCase()
                    let idP1 = (idP[0]) ? (idP[0]) : ''
                    let idP2 = (idP[1]) ? (idP[1]) : ''
                    let idP3 = (idP[2]) ? (idP[2]) : ''
                    let idP4 = (idP[3]) ? (idP[3]) : ''
                    let idP5 = (idP[4]) ? (idP[4]) : ''
                    let idP6 = (idP[5]) ? (idP[5]) : ''
                    let idP7 = (idP[6]) ? (idP[6]) : ''
                    let idP8 = (idP[7]) ? (idP[7]) : ''
                    let idP9 = (idP[8]) ? (idP[8]) : ''
                    let idP10 = (idP[9]) ? (idP[9]) : ''

                    // dateOfIssue {dOI1-8}
                    let dOI = new Date(document.getElementById('dateOfIssue'+indexTab).value).toLocaleDateString().split('.')
                    let dOI1 = (dOI) ? (dOI[0][0]) : ''
                    let dOI2 = (dOI) ? (dOI[0][1]) : ''
                    let dOI3 = (dOI) ? (dOI[1][0]) : ''
                    let dOI4 = (dOI) ? (dOI[1][1]) : ''
                    let dOI5 = (dOI) ? (dOI[2][0]) : ''
                    let dOI6 = (dOI) ? (dOI[2][1]) : ''
                    let dOI7 = (dOI) ? (dOI[2][2]) : ''
                    let dOI8 = (dOI) ? (dOI[2][3]) : ''


                    // validUntil {vU1-8}
                    let vU = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString().split(".") : ''
                    let vU1 = (vU) ? (vU[0][0]) : ''
                    let vU2 = (vU) ? (vU[0][1]) : ''
                    let vU3 = (vU) ? (vU[1][0]) : ''
                    let vU4 = (vU) ? (vU[1][1]) : ''
                    let vU5 = (vU) ? (vU[2][0]) : ''
                    let vU6 = (vU) ? (vU[2][1]) : ''
                    let vU7 = (vU) ? (vU[2][2]) : ''
                    let vU8 = (vU) ? (vU[2][3]) : ''


                    // typeVisa
                    let tVV = ''
                    let tVJ = ''
                    let tVP = ''
                    switch (document.getElementById('typeVisa' + indexTab).value) {
                        case "ВИЗА":
                            tVV = 'X'
                            break
                        case "(ВНЖ) ВИД НА ЖИТЕЛЬСТВО РФ":
                            tVJ = 'X'
                            break
                        case "(РВП) РАЗРЕШЕНИЕ НА ВРЕМЕННОЕ ПРОЖИВАНИЕ РФ":
                            tVP = 'X'
                            break
                    }

                    // seriesVisa {sV1-4}
                    let sV = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value.toUpperCase() : ''
                    let sV1 = ''
                    let sV2 = ''
                    let sV3 = ''
                    let sV4 = ''
                    if (sV) {
                        switch (sV.length) {
                            case 1:
                                sV4 = sV[0]
                                break
                            case 2:
                                sV3 = sV[0]
                                sV4 = sV[1]
                                break
                            case 3:
                                sV2 = sV[0]
                                sV3 = sV[1]
                                sV4 = sV[2]
                                break
                            case 4:
                                sV1 = sV[0]
                                sV2 = sV[1]
                                sV3 = sV[2]
                                sV4 = sV[3]
                        }
                    }

                    // idVisa {idV1-15}
                    let idV = document.getElementById('idVisa' + indexTab).value ? document.getElementById('idVisa' + indexTab).value.toUpperCase() : ''
                    let idV1 = (idV[0]) ? (idV[0]) : ''
                    let idV2 = (idV[1]) ? (idV[1]) : ''
                    let idV3 = (idV[2]) ? (idV[2]) : ''
                    let idV4 = (idV[3]) ? (idV[3]) : ''
                    let idV5 = (idV[4]) ? (idV[4]) : ''
                    let idV6 = (idV[5]) ? (idV[5]) : ''
                    let idV7 = (idV[6]) ? (idV[6]) : ''
                    let idV8 = (idV[7]) ? (idV[7]) : ''
                    let idV9 = (idV[8]) ? (idV[8]) : ''
                    let idV10 = (idV[9]) ? (idV[9]) : ''
                    let idV11 = (idV[10]) ? (idV[10]) : ''
                    let idV12 = (idV[11]) ? (idV[11]) : ''
                    let idV13 = (idV[12]) ? (idV[12]) : ''
                    let idV14 = (idV[13]) ? (idV[13]) : ''
                    let idV15 = (idV[14]) ? (idV[14]) : ''

                    // dateOfIssueVisa {dOIV1-8}
                    let dOIV = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString().split(".") : ''
                    let dOIV1 = (dOIV) ? (dOIV[0][0]) : ''
                    let dOIV2 = (dOIV) ? (dOIV[0][1]) : ''
                    let dOIV3 = (dOIV) ? (dOIV[1][0]) : ''
                    let dOIV4 = (dOIV) ? (dOIV[1][1]) : ''
                    let dOIV5 = (dOIV) ? (dOIV[2][0]) : ''
                    let dOIV6 = (dOIV) ? (dOIV[2][1]) : ''
                    let dOIV7 = (dOIV) ? (dOIV[2][2]) : ''
                    let dOIV8 = (dOIV) ? (dOIV[2][3]) : ''

                    // validUntilVisa {vUV1-8}
                    let vUV = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString().split(".") : ''
                    let vUV1 = (vUV) ? (vUV[0][0]) : ''
                    let vUV2 = (vUV) ? (vUV[0][1]) : ''
                    let vUV3 = (vUV) ? (vUV[1][0]) : ''
                    let vUV4 = (vUV) ? (vUV[1][1]) : ''
                    let vUV5 = (vUV) ? (vUV[2][0]) : ''
                    let vUV6 = (vUV) ? (vUV[2][1]) : ''
                    let vUV7 = (vUV) ? (vUV[2][2]) : ''
                    let vUV8 = (vUV) ? (vUV[2][3]) : ''

                    // purpose
                    let purposeG = ""
                    let purposeR = ""
                    let purposeU = ""
                    let purp1 = ''
                    let purp2 = ''
                    let purp3 = ''
                    let purp4 = ''
                    let purp5 = ''
                    let purp6 = ''
                    let purp7 = ''
                    let purp8 = ''
                    let purp9 = ''
                    let purp10 = ''
                    let purp11 = ''
                    let purp12 = ''
                    let purp13 = ''
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purposeU = "X"
                            purp1 = 'С'; purp2='Т'; purp3='У'; purp4="Д"; purp5 ='Е';purp6 = 'Н';purp7='Т'
                            break
                        case "Краткосрочная учеба":
                            purposeU = "X"
                            purp1 = 'С'; purp2='Т'; purp3='У'; purp4="Д"; purp5 ='Е';purp6 = 'Н';purp7='Т'
                            break
                        case "(НТС)":
                            purposeG = "X"
                            purp1 = 'Н'; purp2='Т'; purp3='С';
                            break
                        case "Трудовая деятельность":
                            purposeR = "X"
                            purp1='П';purp2='Р';purp3='Е';purp4="П";purp5='О';purp6='Д';purp7='А';purp8='В';purp9='А';purp10='Т';purp11='Е';purp12='Л';purp13='Ь';
                            break
                    }

                    // dateArrivalMigration {dAM1-8}
                    let dAM = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString().split(".") : ''
                    let dAM1 = (dAM) ? (dAM[0][0]) : ''
                    let dAM2 = (dAM) ? (dAM[0][1]) : ''
                    let dAM3 = (dAM) ? (dAM[1][0]) : ''
                    let dAM4 = (dAM) ? (dAM[1][1]) : ''
                    let dAM5 = (dAM) ? (dAM[2][0]) : ''
                    let dAM6 = (dAM) ? (dAM[2][1]) : ''
                    let dAM7 = (dAM) ? (dAM[2][2]) : ''
                    let dAM8 = (dAM) ? (dAM[2][3]) : ''

                    // dateUntil {dU1-8}
                    let dU = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString().split(".") : ''
                    let dU1 = (dU) ? (dU[0][0]) : ''
                    let dU2 = (dU) ? (dU[0][1]) : ''
                    let dU3 = (dU) ? (dU[1][0]) : ''
                    let dU4 = (dU) ? (dU[1][1]) : ''
                    let dU5 = (dU) ? (dU[2][0]) : ''
                    let dU6 = (dU) ? (dU[2][1]) : ''
                    let dU7 = (dU) ? (dU[2][2]) : ''
                    let dU8 = (dU) ? (dU[2][3]) : ''

                    // seriesMigration {sM1-4}
                    let sM = document.getElementById('seriesMigration' + indexTab).value
                        ? document.getElementById('seriesMigration' + indexTab).value.toUpperCase() : ''
                    let sM1 = ''
                    let sM2 = ''
                    let sM3 = ''
                    let sM4 = ''
                    if (sM) {
                        switch (sM.length) {
                            case 1:
                                sM4 = sM[0]
                                break
                            case 2:
                                sM3 = sM[0]
                                sM4 = sM[1]
                                break
                            case 3:
                                sM2 = sM[0]
                                sM3 = sM[1]
                                sM4 = sM[2]
                                break
                            case 4:
                                sM1 = sM[0]
                                sM2 = sM[1]
                                sM3 = sM[2]
                                sM4 = sM[3]
                        }
                    }

                    //idMigration {iM1-11}
                    let iM = document.getElementById('idMigration' + indexTab).value
                        ? document.getElementById('idMigration' + indexTab).value.toUpperCase() : ''
                    let iM1 = (iM[0]) ? (iM[0]) : ''
                    let iM2 = (iM[1]) ? (iM[1]) : ''
                    let iM3 = (iM[2]) ? (iM[2]) : ''
                    let iM4 = (iM[3]) ? (iM[3]) : ''
                    let iM5 = (iM[4]) ? (iM[4]) : ''
                    let iM6 = (iM[5]) ? (iM[5]) : ''
                    let iM7 = (iM[6]) ? (iM[6]) : ''
                    let iM8 = (iM[7]) ? (iM[7]) : ''
                    let iM9 = (iM[8]) ? (iM[8]) : ''
                    let iM10 = (iM[9]) ? (iM[9]) : ''
                    let iM11 = (iM[10]) ? (iM[10]) : ''

                    // addressHostel {aHG1-18} {aHU1-20} {aHD1-15} {aHK1-5}
                    let aHG1 = ''
                    let aHG2 = ''
                    let aHG3 = ''
                    let aHG4 = ''
                    let aHG5 = ''
                    let aHG6 = ''
                    let aHG7 = ''
                    let aHG8 = ''
                    let aHG9 = ''
                    let aHG10 = ''
                    let aHG11 = ''
                    let aHG12 = ''
                    let aHG13 = ''
                    let aHG14 = ''
                    let aHG15 = ''
                    let aHG16 = ''
                    let aHG17 = ''
                    let aHG18 = ''
                    let aHU1 = ''
                    let aHU2 = ''
                    let aHU3 = ''
                    let aHU4 = ''
                    let aHU5 = ''
                    let aHU6 = ''
                    let aHU7 = ''
                    let aHU8 = ''
                    let aHU9 = ''
                    let aHU10 = ''
                    let aHU11 = ''
                    let aHU12 = ''
                    let aHU13 = ''
                    let aHU14 = ''
                    let aHU15 = ''
                    let aHU16 = ''
                    let aHU17 = ''
                    let aHU18 = ''
                    let aHU19 = ''
                    let aHU20 = ''
                    let aHD1 = ""
                    let aHD2 = ""
                    let aHD3 = ""
                    let aHD4 = ""
                    let aHD5 = ""
                    let aHD6 = ""
                    let aHD7 = ""
                    let aHD8 = ""
                    let aHD9 = ""
                    let aHD10 = ""
                    let aHD11 = ""
                    let aHD12 = ""
                    let aHD13 = ""
                    let aHD14 = ""
                    let aHD15 = ""
                    let aHK1 = ''
                    let aHK2 = ''
                    let aHK3 = ''
                    let aHK4 = ''
                    let aHK5 = ''

                    switch (document.getElementById('addressHostel'+indexTab).value) {
                        case "г. Москва, проспект Вернадского, 88 к. 1 (ОБЩЕЖИТИЕ №1)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '1';
                            break
                        case "г. Москва, проспект Вернадского, 88 к. 2 (ОБЩЕЖИТИЕ №2)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '2';
                            break
                        case "г. Москва, проспект Вернадского, 88 к. 3 (ОБЩЕЖИТИЕ №3)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '3';
                            break
                        case "г. Москва, улица Космонавтов, д. 13 (ОБЩЕЖИТИЕ №4)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'К'; aHU2 = 'О'; aHU3 = 'С'; aHU4 = 'М'; aHU5 = 'О'; aHU6 = 'Н'; aHU7 = 'А'; aHU8 = 'В'; aHU9 = 'Т'; aHU10 = 'О'; aHU11 = 'В';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '1'; aHD6 = '3';
                            break
                        case "г. Москва, улица Космонавтов, д. 9 (ОБЩЕЖИТИЕ №5)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'К'; aHU2 = 'О'; aHU3 = 'С'; aHU4 = 'М'; aHU5 = 'О'; aHU6 = 'Н'; aHU7 = 'А'; aHU8 = 'В'; aHU9 = 'Т'; aHU10 = 'О'; aHU11 = 'В';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '9';
                            break
                        case "г. Москва, улица Клары Цеткин, д. 25 (ОБЩЕЖИТИЕ №6)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'К'; aHU2 = 'Л'; aHU3 = 'А'; aHU4 = 'Р'; aHU5 = 'Ы'; aHU6 = ''; aHU7 = 'Ц'; aHU8 = 'Е'; aHU9 = 'Т'; aHU10 = 'К'; aHU11 = 'И'; aHU12 = 'Н';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '2'; aHD6 = '5';
                            break
                        case "Московская область, г. Люберцы, ул. Мира, д.7 (ОБЩЕЖИТИЕ №7)":
                            aHG1 = 'М'; aHG2 = 'О'; aHG3 = 'С'; aHG4 = 'К'; aHG5 = 'О'; aHG6 = 'В'; aHG7 = 'С'; aHG8 = 'К'; aHG9 = 'А'; aHG10 = 'Я'; aHG12 = 'О'; aHG13 = 'Б'; aHG14 = 'Л'; aHG15 = 'А'; aHG16 = 'С'; aHG17 = 'Т'; aHG18 = 'Ь';
                            aHU1 = 'Л'; aHU2 = 'Ю'; aHU3 = 'Б'; aHU4 = 'Е'; aHU5 = 'Р'; aHU6 = 'Ц'; aHU7 = 'Ы';
                            aHD1 = 'М'; aHD2 = 'И'; aHD3 = 'Р'; aHD4 = 'А';
                            aHK1 = 'Д'; aHK2 = 'О'; aHK3 = 'М'; aHK4 = ''; aHK5 = '7';
                            break
                    }

                    // migrationAddress {mAO1-18} {mAG1-7} {mAU1-20} {mAD1-2} {mAK1}
                    let mAO1 = ""
                    let mAO2 = ""
                    let mAO3 = ""
                    let mAO4 = ""
                    let mAO5 = ""
                    let mAO6 = ""
                    let mAO7 = ""
                    let mAO8 = ""
                    let mAO9 = ""
                    let mAO10 = ""
                    let mAO11 = ""
                    let mAO12 = ""
                    let mAO13 = ""
                    let mAO14 = ""
                    let mAO15 = ""
                    let mAO16 = ""
                    let mAO17 = ""
                    let mAO18 = ""
                    let mAG1 = ''
                    let mAG2 = ''
                    let mAG3 = ''
                    let mAG4 = ''
                    let mAG5 = ''
                    let mAG6 = ''
                    let mAG7 = ''
                    let mAU1 = ""
                    let mAU2 = ""
                    let mAU3 = ""
                    let mAU4 = ""
                    let mAU5 = ""
                    let mAU6 = ""
                    let mAU7 = ""
                    let mAU8 = ""
                    let mAU9 = ""
                    let mAU10 = ""
                    let mAU11 = ""
                    let mAU12 = ""
                    let mAU13 = ""
                    let mAU14 = ""
                    let mAU15 = ""
                    let mAU16 = ""
                    let mAU17 = ""
                    let mAU18 = ""
                    let mAU19 = ""
                    let mAU20 = ""
                    let mAD1 = ''
                    let mAD2 = ''
                    let mAK1 = ''
                    switch (document.getElementById('migrationAddress').value) {
                        case 'г. Москва, проспект Вернадского, 88 к. 1':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                            mAD1 = '8'; mAD2 = '8'
                            mAK1 = '1'
                            break
                        case 'г. Москва, проспект Вернадского, 88 к. 2':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                            mAD1 = '8'; mAD2 = '8'
                            mAK1 = '2'
                            break
                        case 'г. Москва, проспект Вернадского, 88 к. 3':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                            mAD1 = '8'; mAD2 = '8'
                            mAK1 = '3'
                            break
                        case 'г. Москва, ул. Космонавтов, д. 13':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'К'; mAU2 = 'О'; mAU3 = 'С'; mAU4 = 'М'; mAU5 = 'О'; mAU6 = 'Н'; mAU7 = 'А'; mAU8 = 'В'; mAU9 = 'Т'; mAU10 = 'О'; mAU11 = 'В';
                            mAD1 = '1'; mAD2 = '3'
                            break
                        case 'г. Москва, улица Космонавтов, д. 9':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'К'; mAU2 = 'О'; mAU3 = 'С'; mAU4 = 'М'; mAU5 = 'О'; mAU6 = 'Н'; mAU7 = 'А'; mAU8 = 'В'; mAU9 = 'Т'; mAU10 = 'О'; mAU11 = 'В';
                            mAD2 = '9'
                            break
                        case 'г. Москва, ул. Клары Цеткин, д. 25, корп. 1':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'К'; mAU2 = 'Л'; mAU3 = 'А'; mAU4 = 'Р'; mAU5 = 'Ы'; mAU7 = 'Ц'; mAU8 = 'Е'; mAU9 = 'Т'; mAU10 = 'К'; mAU11 = 'И'; mAU12 = 'Н';
                            mAD1 = '2'; mAD2 = '5'
                            mAK1 = '1'
                            break
                        case 'Московская область, г. Люберцы, ул. Мира, д.7':
                            mAO1 = 'М'; mAO2 = 'О'; mAO3 = 'С'; mAO4 = 'К'; mAO5 = 'О'; mAO6 = 'В'; mAO7 = 'С'; mAO8 = 'К'; mAO9 = 'А'; mAO10 = 'Я'; mAO12 = 'О'; mAO13 = 'Б'; mAO14 = 'Л'; mAO15 = 'А'; mAO16 = 'С'; mAO17 = 'Т'; mAO18 = 'Ь';
                            mAG1 = 'Л'; mAG2 = 'Ю'; mAG3 = 'Б'; mAG4 = 'Е'; mAG5 = 'Р'; mAG6 = 'Ц'; mAG7 = 'Ы';
                            mAD2 = '7'
                            break
                    }


                    //numRoom {nR1-4}
                    let nR = document.getElementById('numRoom'+ indexTab).value != "-"
                        ? document.getElementById('numRoom' + indexTab).value.toUpperCase() : ''
                    let nR1 = ''
                    let nR2 = ''
                    let nR3 = ''
                    let nR4 = ''
                    if (nR) {
                        switch (nR.length) {
                            case 1:
                                nR4 = nR[0]
                                break
                            case 2:
                                nR3 = nR[0]
                                nR4 = nR[1]
                                break
                            case 3:
                                nR2 = nR[0]
                                nR3 = nR[1]
                                nR4 = nR[2]
                                break
                            case 4:
                                nR1 = nR[0]
                                nR2 = nR[1]
                                nR3 = nR[2]
                                nR4 = nR[3]
                        }
                    }






                    doc.setData({
                        lN1: lN1, lN2: lN2, lN3: lN3, lN4: lN4, lN5: lN5, lN6: lN6, lN7: lN7, lN8: lN8, lN9: lN9, lN10: lN10, lN11: lN11, lN12: lN12, lN13: lN13, lN14: lN14, lN15: lN15, lN16: lN16, lN17: lN17, lN18: lN18, lN19: lN19, lN20: lN20, lN21: lN21, lN22: lN22, lN23: lN23, lN24: lN24, lN25: lN25, lN26: lN26, lN27: lN27,
                        fN1: fN1, fN2: fN2, fN3: fN3, fN4: fN4, fN5: fN5, fN6: fN6, fN7: fN7, fN8: fN8, fN9: fN9, fN10: fN10, fN11: fN11, fN12: fN12, fN13: fN13, fN14: fN14, fN15: fN15, fN16: fN16, fN17: fN17, fN18: fN18, fN19: fN19, fN20: fN20, fN21: fN21, fN22: fN22, fN23: fN23, fN24: fN24, fN25: fN25, fN26: fN26, fN27: fN27,
                        patr1: patr1, patr2: patr2, patr3: patr3, patr4: patr4, patr5: patr5, patr6: patr6, patr7: patr7, patr8: patr8, patr9: patr9, patr10: patr10, patr11: patr11, patr12: patr12, patr13: patr13, patr14: patr14, patr15: patr15, patr16: patr16, patr17: patr17, patr18: patr18, patr19: patr19, patr20: patr20, patr21: patr21, patr22: patr22, patr23: patr23, patr24: patr24,
                        grazd1: grazd1, grazd2: grazd2, grazd3: grazd3, grazd4: grazd4, grazd5: grazd5, grazd6: grazd6, grazd7: grazd7, grazd8: grazd8, grazd9: grazd9, grazd10: grazd10, grazd11: grazd11, grazd12: grazd12, grazd13: grazd13, grazd14: grazd14, grazd15: grazd15, grazd16: grazd16, grazd17: grazd17, grazd18: grazd18, grazd19: grazd19, grazd20: grazd20, grazd21: grazd21, grazd22: grazd22, grazd23: grazd23, grazd24: grazd24, grazd25: grazd25,
                        dOB1: dOB1, dOB2: dOB2, dOB3: dOB3, dOB4: dOB4, dOB5: dOB5, dOB6: dOB6, dOB7: dOB7, dOB8: dOB8,
                        genM: genM, genW: genW,
                        pSBS1: pSBS1, pSBS2: pSBS2, pSBS3: pSBS3, pSBS4: pSBS4, pSBS5: pSBS5, pSBS6: pSBS6, pSBS7: pSBS7, pSBS8: pSBS8, pSBS9: pSBS9, pSBS10: pSBS10, pSBS11: pSBS11, pSBS12: pSBS12, pSBS13: pSBS13, pSBS14: pSBS14, pSBS15: pSBS15, pSBS16: pSBS16, pSBS17: pSBS17, pSBS18: pSBS18, pSBS19: pSBS19, pSBS20: pSBS20, pSBS21: pSBS21, pSBS22: pSBS22, pSBS23: pSBS23, pSBS24: pSBS24,
                        pSBG1: pSBG1, pSBG2: pSBG2, pSBG3: pSBG3, pSBG4: pSBG4, pSBG5: pSBG5, pSBG6: pSBG6, pSBG7: pSBG7, pSBG8: pSBG8, pSBG9: pSBG9, pSBG10: pSBG10, pSBG11: pSBG11, pSBG12: pSBG12, pSBG13: pSBG13, pSBG14: pSBG14, pSBG15: pSBG15, pSBG16: pSBG16, pSBG17: pSBG17, pSBG18: pSBG18, pSBG19: pSBG19, pSBG20: pSBG20, pSBG21: pSBG21, pSBG22: pSBG22, pSBG23: pSBG23, pSBG24: pSBG24,
                        sP1: sP1, sP2: sP2, sP3: sP3, sP4: sP4,
                        idP1: idP1, idP2: idP2, idP3: idP3, idP4: idP4, idP5: idP5, idP6: idP6, idP7: idP7, idP8: idP8, idP9: idP9, idP10: idP10,
                        dOI1: dOI1, dOI2: dOI2, dOI3: dOI3, dOI4: dOI4, dOI5: dOI5, dOI6: dOI6, dOI7: dOI7, dOI8: dOI8,
                        vU1: vU1, vU2: vU2, vU3: vU3, vU4: vU4, vU5: vU5, vU6: vU6, vU7: vU7, vU8: vU8,
                        tVV: tVV, tVJ: tVJ, tVP: tVP,
                        sV1: sV1, sV2: sV2, sV3: sV3, sV4: sV4,
                        idV1: idV1, idV2: idV2, idV3: idV3, idV4: idV4, idV5: idV5, idV6: idV6, idV7: idV7, idV8: idV8, idV9: idV9, idV10: idV10, idV11: idV11, idV12: idV12, idV13: idV13, idV14: idV14, idV15: idV15,
                        dOIV1: dOIV1, dOIV2: dOIV2, dOIV3: dOIV3, dOIV4: dOIV4, dOIV5: dOIV5, dOIV6: dOIV6, dOIV7: dOIV7, dOIV8: dOIV8,
                        vUV1: vUV1, vUV2: vUV2, vUV3: vUV3, vUV4: vUV4, vUV5: vUV5, vUV6: vUV6, vUV7: vUV7, vUV8: vUV8,
                        dAM1: dAM1, dAM2: dAM2, dAM3: dAM3, dAM4: dAM4, dAM5: dAM5, dAM6: dAM6, dAM7: dAM7, dAM8: dAM8,
                        dU1: dU1, dU2: dU2, dU3: dU3, dU4: dU4, dU5: dU5, dU6: dU6, dU7: dU7, dU8: dU8,
                        sM1: sM1, sM2: sM2, sM3: sM3, sM4: sM4,
                        iM1: iM1, iM2: iM2, iM3: iM3, iM4: iM4, iM5: iM5, iM6: iM6, iM7: iM7, iM8: iM8, iM9: iM9, iM10: iM10, iM11: iM11,
                        aHG1: aHG1, aHG2: aHG2, aHG3: aHG3, aHG4: aHG4, aHG5: aHG5, aHG6: aHG6, aHG7: aHG7, aHG8: aHG8, aHG9: aHG9, aHG10: aHG10, aHG11: aHG11, aHG12: aHG12, aHG13: aHG13, aHG14: aHG14, aHG15: aHG15, aHG16: aHG16, aHG17: aHG17, aHG18: aHG18,
                        aHU1: aHU1, aHU2: aHU2, aHU3: aHU3, aHU4: aHU4, aHU5: aHU5, aHU6: aHU6, aHU7: aHU7, aHU8: aHU8, aHU9: aHU9, aHU10: aHU10, aHU11: aHU11, aHU12: aHU12, aHU13: aHU13, aHU14: aHU14, aHU15: aHU15, aHU16: aHU16, aHU17: aHU17, aHU18: aHU18, aHU19: aHU19, aHU20: aHU20,
                        aHD1: aHD1, aHD2: aHD2, aHD3: aHD3, aHD4: aHD4, aHD5: aHD5, aHD6: aHD6, aHD7: aHD7, aHD8: aHD8, aHD9: aHD9, aHD10: aHD10, aHD11: aHD11, aHD12: aHD12, aHD13: aHD13, aHD14: aHD14, aHD15: aHD15,
                        aHK1: aHK1, aHK2: aHK2, aHK3: aHK3, aHK4: aHK4, aHK5: aHK5,
                        mAO1: mAO1, mAO2: mAO2, mAO3: mAO3, mAO4: mAO4, mAO5: mAO5, mAO6: mAO6, mAO7: mAO7, mAO8: mAO8, mAO9: mAO9, mAO10: mAO10, mAO11: mAO11, mAO12: mAO12, mAO13: mAO13, mAO14: mAO14, mAO15: mAO15, mAO16: mAO16, mAO17: mAO17, mAO18: mAO18,
                        mAG1: mAG1, mAG2: mAG2, mAG3: mAG3, mAG4: mAG4, mAG5: mAG5, mAG6: mAG6, mAG7: mAG7,
                        mAU1: mAU1, mAU2: mAU2, mAU3: mAU3, mAU4: mAU4, mAU5: mAU5, mAU6: mAU6, mAU7: mAU7, mAU8: mAU8, mAU9: mAU9, mAU10: mAU10, mAU11: mAU11, mAU12: mAU12, mAU13: mAU13, mAU14: mAU14, mAU15: mAU15, mAU16: mAU16, mAU17: mAU17, mAU18: mAU18, mAU19: mAU19, mAU20: mAU20,
                        mAD1: mAD1, mAD2: mAD2, mAK1: mAK1,


                        purp1: purp1, purp2: purp2, purp3: purp3, purp4: purp4, purp5: purp5, purp6: purp6, purp7: purp7, purp8: purp8, purp9: purp9, purp10: purp10, purp11: purp11, purp12: purp12, purp13: purp13,
                        purposeG: purposeG, purposeU: purposeU, purposeR: purposeR,

                        nR1: nR1, nR2: nR2, nR3: nR3, nR4: nR4,


                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(

                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + "/"+
                        "УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         +".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         + ".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                // saveAs(content,nameFile);
            });
    };

//ходатайство РЕГИСТРАЦИЯ
    window.generateRegSolicitaionTotal = function generate() {
        path = ("")
        switch (document.getElementById('ovmByRegion').value) {
            case "Алексеевский":
                path = ('../Templates/регистрация/ходатайство АЛЕКСЕЕВСКИЙ.docx')
                break
            case "Войковский":
                path = ('../Templates/регистрация/ходатайство ВОЙКОВСКИЙ.docx')
                break
            case "МУ МВД РФ Люберецкое":
                path = ('../Templates/регистрация/ходатайство МУ МВД РФ ЛЮБЕРЕЦКОЕ.docx')
                break
            case "Тропарево-Никулино":
                path = ('../Templates/регистрация/ходатайство ТРОПАРЕВО-НИКУЛИНО.docx')
                break
        }

        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))


                    //dateUntil
                    let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                    // purpose
                    let purpose = ''
                    let purposeS = ''
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purpose = "УЧЕБА"
                            purposeS = "Студент"
                            break
                        case "Краткосрочная учеба":
                            purpose = "КРАТКОСРОЧНАЯ УЧЕБА"
                            purposeS = "Студент"
                            break
                        case "(НТС)":
                            purpose = "НАУЧНО-ТЕХНИЧЕСКИЕ СВЯЗИ (НТС)"
                            purposeS = "НТС"
                            break
                        case "Трудовая деятельность":
                            purpose = "ТРУДОВАЯ ДЕЯТЕЛЬНОСТЬ"
                            purposeS = "Преподаватель"
                            break
                    }

                    // levelEducation
                    let levelEducation = ''
                    switch (document.getElementById('levelEducation' + indexTab).value) {
                        case "Подготовительный факультет (изучаю русский язык)/ The preparatory faculty":
                            levelEducation = 'подготовительный факультет'
                            break
                        case "бакалавриат/bachelor degree":
                            levelEducation = 'бакалавриат'
                            break
                        case "магистратура/master degree":
                            levelEducation = 'магистратура'
                            break
                        case "аспирантура/post-graduate studies":
                            levelEducation = 'аспирантура'
                            break
                    }

                    // course
                    let course = ''
                    switch (document.getElementById('course' + indexTab).value) {
                        case '1':
                            course = ', 1 курс,'
                            break
                        case '2':
                            course = ', 2 курс,'
                            break
                        case '3':
                            course = ', 3 курс,'
                            break
                        case '4':
                            course = ', 4 курс,'
                            break
                        case '5':
                            course = ', 5 курс,'
                            break
                    }

                    // addressResidence
                    let addressResidence = ''
                    switch (document.getElementById('migrationAddress').value) {
                        case "Квартира":
                            addressResidence = document.getElementById('addressResidence' + indexTab).value
                            break
                        default:
                            addressResidence = document.getElementById('migrationAddress').value
                            break
                    }

                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                    // gender
                    let gender = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            gender = 'м.'
                            break
                        case "Женский / Female":
                            gender = 'ж.'
                            break
                    }

                    // registration On
                    let registrationOn = ''
                    switch (document.getElementById('registrationOn').value) {
                        case "Круглов":
                            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                            break
                        case "Морозова":
                            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                            break
                        case "Орлова":
                            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                            break
                    }

                    // visa
                    let typeVisa = ''
                    switch (document.getElementById('typeVisa' + indexTab).value) {
                        case "ВИЗА":
                            typeVisa = 'Виза'
                            break
                        case "(ВНЖ) ВИД НА ЖИТЕЛЬСТВО РФ":
                            typeVisa = 'ВНЖ'
                            break
                        case "(РВП) РАЗРЕШЕНИЕ НА ВРЕМЕННОЕ ПРОЖИВАНИЕ РФ":
                            typeVisa = 'РВП'
                            break

                    }
                    let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value : ''
                    let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                        ? document.getElementById('idVisa' + indexTab).value : ''

                    let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                    let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                    // order
                    let numOrder = document.getElementById('numOrder' + indexTab).value
                        ? document.getElementById('numOrder' + indexTab).value : ''
                    let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''
                    let orderUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString() : ''

                    // faculty
                    let faculty = ''
                    switch (document.getElementById('faculty' + indexTab).value) {
                        case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                            faculty = 'ИИИ:Музфак'
                            break
                        case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                            faculty = 'ИИИ: Худграф'
                            break
                        case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                            faculty = 'ИСГО'
                            break
                        case "Институт филологии / The Institute of Philology":
                            faculty = 'ИФ'
                            break
                        case "Институт иностранных языков / The Institute of Foreign Languages":
                            faculty = 'ИИЯ'
                            break
                        case "Институт международного образования / The Institute of International Education":
                            faculty = 'ИМО'
                            break
                        case "Институт детства / The Institute of Childhood":
                            faculty = 'ИД'
                            break
                        case "Институт биологии и химии / The Institute of Biology and Chemistry":
                            faculty = 'ИБХ'
                            break
                        case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                            faculty = 'ИФТИС'
                            break
                        case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                            faculty = 'ИФКСиЗ'
                            break
                        case "Географический факультет / The Institute of Geography":
                            faculty = 'Геофак'
                            break
                        case "Институт истории и политики / The Institute of History and Politics":
                            faculty = 'ИИП'
                            break
                        case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                            faculty = 'ИМИ'
                            break
                        case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                            faculty = 'Дош.фак.'
                            break
                        case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                            faculty = 'ИПП'
                            break
                        case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                            faculty = 'ИЖКиМ'
                            break
                        case "Институт развития цифрового образования / The Institute of Digital Education Development":
                            faculty = 'ИРЦО'
                            break
                    }


                    // contract
                    let typeFundingDog1 = ""
                    let typeFundingDog2 = ""
                    let typeFundingNap1 = ""
                    let typeFundingNap2 = ""
                    switch (document.getElementById('typeFunding' + indexTab).value) {
                        case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                            typeFundingDog2 = "Договор"
                            typeFundingNap2 = "направление"
                            break
                        case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                            typeFundingDog1 = "Договор"
                            typeFundingNap1 = "направление"
                            break
                    }
                    let numContract = document.getElementById('numContract' + indexTab).value
                        ? document.getElementById('numContract' + indexTab).value : ''
                    let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                        document.getElementById('contractFrom' + indexTab).value : ''


                    let numRoom = document.getElementById('numRoom' + indexTab).value != "" ? ', комната № ' + document.getElementById('numRoom' + indexTab).value : ''

                    let numRental = document.getElementById('numRental' + indexTab).value != "-" ? document.getElementById('numRental' + indexTab).value : ''


                    doc.setData({
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                        nStud: document.getElementById('nStud' + indexTab).value,
                        grazd: document.getElementById('grazd' + indexTab).value,
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                        dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                        gender: gender,

                        registrationOn: registrationOn,
                        dateUntil: dateUntil,
                        purpose: purpose,
                        purposeS: purposeS,
                        levelEducation: levelEducation,
                        course: course,

                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,

                        typeVisa: typeVisa,
                        seriesVisa: seriesVisa,
                        idVisa: idVisa,
                        dateOfIssueVisa: dateOfIssueVisa,
                        validUntilVisa: validUntilVisa,

                        seriesMigration: document.getElementById('seriesMigration' + indexTab).value,
                        idMigration: document.getElementById('idMigration' + indexTab).value,
                        dateArrivalMigration: new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString(),



                        migrationAddress: document.getElementById('migrationAddress').value,
                        numRoom: '',
                        faculty: faculty,
                        numOrder: numOrder,
                        orderFrom: orderFrom,
                        orderUntil: orderUntil,

                        typeFundingDog1: typeFundingDog1,
                        typeFundingDog2: typeFundingDog2,
                        typeFundingNap1: typeFundingNap1,
                        typeFundingNap2: typeFundingNap2,

                        numContract: numContract,
                        contractFrom: contractFrom,

                        numRental: '',

                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + "/" +

                        "ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         +".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         + ".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                //saveAs(content,nameFile);
            });
    };

//опись РЕГИСТРАЦИЯ
    window.generateInventoryRegTotal = function generate() {
        path = ('../Templates/регистрация/опись регистрация.docx')

        let students = []

        // registration On
        let registrationOn = ''
        switch (document.getElementById('registrationOn').value) {
            case "Круглов":
                registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                break
            case "Морозова":
                registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                break
            case "Орлова":
                registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                break
        }

        // ovmByRegion
        let ovmByRegion = ''
        switch (document.getElementById('ovmByRegion').value) {
            case 'Тропарево-Никулино':
                ovmByRegion = 'ОМВД России по району Тропарево-Никулино г. Москвы'
                break
            case 'Хамовники':
                ovmByRegion = 'ОМВД России по району Хамовники г. Москвы'
                break
            case 'Алексеевский':
                ovmByRegion = 'ОМВД России по Алексеевскому району г.Москвы'
                break
            case 'Войковский':
                ovmByRegion = 'ОВМ МВД России по району Войковский г. Москвы'
                break
            case 'МУ МВД РФ Люберецкое':
                ovmByRegion = 'ОВМ МУ МВД РФ "Люберецкое"'
                break
        }

        // nStud
        let nStud1 = document.getElementById('nStud1').value
        let tir = ''
        let nStud2 = ''
        if (countTab()>1) {
            nStud2 = document.getElementById('nStud' + (lastTab()-1)).value
            tir = '-'
        }

        for (let i =0; i<countTab();i++) {
            let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
            let elem = tabs[i]
            let indexTab = parseInt(elem.id.match(/\d+/))



            let dateUntil = document.getElementById('dateUntil' + indexTab).value != '-' ?
                new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

            //students
            students.push({
                nStud: document.getElementById('nStud' + indexTab).value,
                lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                firstNameRu: document.getElementById('firstNameRu'+indexTab).value.toUpperCase(),
                patronymicRu: document.getElementById('patronymicRu'+indexTab).value.toUpperCase(),
                dateInOv: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                dateUntil: dateUntil,
                grazd: document.getElementById('grazd' + indexTab).value,
                phone: document.getElementById('phone' + indexTab).value,
                mail: document.getElementById('mail' + indexTab).value,
            })
        }


        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                doc.render({
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    nStud1: nStud1,
                    nStud2: nStud2,
                    tir: tir,
                    ovmByRegion: ovmByRegion,
                    'students': students,
                    registrationOn: registrationOn,
                })


                var out = doc.getZip().generate();



                let nameFile = ''
                if (countTab()==1) {
                    nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ) - студент " +
                        document.getElementById('nStud1').value
                        +" - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                        +" - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                }
                else {
                    nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ) - студенты " +
                        document.getElementById('nStud1').value
                        +"-"+
                        document.getElementById('nStud'+(lastTab()-1)).value
                        +" - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                        +" - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                }

                zipTotal.file(nameFile, out, {base64: true})

                // Output the document using Data-URI
                //saveAs(out, nameFile);
                var content = zipTotal.generate({type: 'blob'})

                let nameZip = ''
                if (countTab()==1) {
                    nameZip = document.getElementById('nStud1').value
                        +" Студент.zip"
                }
                else {
                    nameZip = document.getElementById('nStud1').value + '-'+
                        document.getElementById('nStud'+(lastTab()-1)).value
                        +" Студенты.zip"
                }

                saveAs(content, nameZip)
            }
        );
    };




    setTimeout(generateRegNotifTotal, 30)
    setTimeout(generateRegSolicitaionTotal, 30)
    setTimeout(generateInventoryRegTotal, 300)
}





// ВИЗА общий выгруз

function generateVisa() {
    var zipTotal = new PizZip();

//визовая анкета
    window.generateVisaApplicationTotal = function generate() {
        path = ('../Templates/виза/визовая анкета.docx')
        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))

                    // purpose
                    let purposeG = ""
                    let purposeR = ""
                    let purposeU = ""
                    let purposeS = "-"
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purposeU = "X"
                            purposeS = "Студент"
                            break
                        case "Краткосрочная учеба":
                            purposeU = "X"
                            purposeS = "Студент"
                            break
                        case "(НТС)":
                            purposeG = "X"
                            break
                        case "Трудовая деятельность":
                            purposeR = "X"
                            purposeS = "Преподаватель"
                            break
                    }

                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                    // gender
                    let genM = ''
                    let genW = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            genM = 'X'
                            break
                        case "Женский / Female":
                            genW = 'X'
                            break
                    }

                    // visa
                    let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value : ''
                    let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                        ? document.getElementById('idVisa' + indexTab).value : ''
                    let identifierVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('identifierVisa' + indexTab).value)
                        ? document.getElementById('identifierVisa' + indexTab).value : ''
                    let numInvVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('numInvVisa' + indexTab).value)
                        ? document.getElementById('numInvVisa' + indexTab).value : ''

                    let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                    let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''


                    // OVM
                    let infHost1 = ''
                    let infHost2 = ''
                    let addressResidence = ''
                    let numRoom = ''
                    switch (document.getElementById('migrationAddress').value) {
                        case "Квартира":
                            if (document.getElementById('infHost' + indexTab).value.length>60) {
                                let totalLen = 0
                                let lenInfHost = document.getElementById('infHost' + indexTab).value.split(',')
                                for (let i = 0; i< lenInfHost.length; i++) {
                                    totalLen = totalLen + lenInfHost[i].length
                                    if (totalLen<60) {
                                        if (i==lenInfHost.length-1) {infHost1 = infHost1  + lenInfHost[i]}
                                        else {infHost1 = infHost1  + lenInfHost[i] + ', '}
                                    }
                                    else {
                                        if (i==lenInfHost.length-1) {infHost2 = infHost2  + lenInfHost[i]}
                                        else {infHost2 = infHost2  + lenInfHost[i] + ', '}
                                    }
                                }
                            }
                            else {infHost1 = document.getElementById('infHost' + indexTab).value}
                            addressResidence = document.getElementById('addressResidence' + indexTab).value
                            break
                        default:
                            infHost1 = 'МПГУ, 7704077771, М. Пироговская д. 1, стр. 1.'
                            infHost2 = '8-499-245-03-10, mail@mpgu.su'
                            addressResidence = document.getElementById('migrationAddress').value
                            numRoom = ', комната № ' + document.getElementById('numRoom' + indexTab).value
                            break

                    }

                    doc.setData({


                        purposeG: purposeG,
                        purposeR: purposeR,
                        purposeU: purposeU,
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                        lastNameEn: document.getElementById('lastNameEn' + indexTab).value.toUpperCase(),
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value.toUpperCase(),
                        firstNameEn: document.getElementById('firstNameEn' + indexTab).value.toUpperCase(),
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value.toUpperCase(),
                        dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                        placeStateBirth: document.getElementById('placeStateBirth' + indexTab).value.toUpperCase(),
                        genM: genM,
                        genW: genW,
                        grazd: document.getElementById('grazd' + indexTab).value,
                        // documentPerson: document.getElementById('documentPerson' + indexTab).text,
                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,

                        infHost1: infHost1,
                        infHost2: infHost2,
                        addressResidence: addressResidence,
                        numRoom: '',

                        homeAddress: document.getElementById('homeAddress' + indexTab).value,
                        purposeS: purposeS,
                        phone: document.getElementById('phone' + indexTab).value,
                        mail: document.getElementById('mail' + indexTab).value,
                        seriesVisa: seriesVisa,
                        idVisa: idVisa,
                        identifierVisa: identifierVisa,
                        dateOfIssueVisa: dateOfIssueVisa,
                        validUntilVisa: validUntilVisa,
                        numInvVisa: numInvVisa,
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),

                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + "/" +

                        "ВИЗОВАЯ АНКЕТА - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" ВИЗОВАЯ АНКЕТА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text +".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" ВИЗОВАЯ АНКЕТА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text +".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                //saveAs(content,nameFile);
            });
    };

//ходатайство ВИЗА ТРОПАРЕВО-НИКУЛИНО
    window.generateVisaSolicitaionTroparevoTotal = function generate() {
        path = ('../Templates/виза/ходатайство ТРОПАРЕВО-НИКУЛИНО.docx')

        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))


                    //dateUntil
                    let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                    // purpose
                    let purpose = ''
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purpose = "обучением в МПГУ"
                            break
                        case "Краткосрочная учеба":
                            purpose = "обучением в МПГУ"
                            break
                        case "(НТС)":
                            purpose = "посещением МПГУ в качестве приглашенного гостя (НТС)"
                            break
                        case "Трудовая деятельность":
                            purpose = "преподавательской деятельностью в МПГУ"
                            break
                    }

                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                    // gender
                    let genM = ''
                    let genW = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            genM = 'X'
                            genW = ' '
                            break
                        case "Женский / Female":
                            genW = 'X'
                            genM = ' '
                            break
                    }

                    // registration On
                    let registrationOn1 = ''
                    let registrationOn2 = ''
                    switch (document.getElementById('registrationOn').value) {
                        case "Круглов":
                            registrationOn1 = 'Начальник УМС                                                    Круглов В.В.'
                            registrationOn2 = 'Начальник УМС                                                                          В. В. Круглов'
                            break
                        case "Морозова":
                            registrationOn1 = 'Заместитель начальника УМС                Морозова О.А.'
                            registrationOn2 = 'Заместитель начальника УМС                                                Морозова О.А.'
                            break
                        case "Орлова":
                            registrationOn1 = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                            registrationOn2 = 'Начальник паспортно-визового отдела УМС                              Орлова С.В.'
                            break
                    }

                    // visa
                    let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value : ''
                    let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                        ? document.getElementById('idVisa' + indexTab).value : ''

                    let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                    let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                    // order
                    let numOrder = document.getElementById('numOrder' + indexTab).value
                        ? document.getElementById('numOrder' + indexTab).value : ''
                    let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''
                    let orderUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString() : ''

                    // contract
                    let typeFunding = ''
                    switch (document.getElementById('typeFunding' + indexTab).value) {
                        case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                            typeFunding = 'НАПРАВЛЕНИЕ №'
                            break
                        case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                            typeFunding = 'ДОГОВОР №'
                            break
                    }
                    let numContract = document.getElementById('numContract' + indexTab).value
                        ? document.getElementById('numContract' + indexTab).value : ''
                    let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                        document.getElementById('contractFrom' + indexTab).value : ''



                    doc.setData({
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                        nStud: document.getElementById('nStud' + indexTab).value,
                        grazd: document.getElementById('grazd' + indexTab).value,
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                        lastNameEn: document.getElementById('lastNameEn' + indexTab).value,
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                        firstNameEn: document.getElementById('firstNameEn' + indexTab).value,
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                        dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                        registrationOn1: registrationOn1,
                        registrationOn2: registrationOn2,
                        dateUntil: dateUntil,
                        genM: genM,
                        genW: genW,
                        purpose: purpose,

                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,

                        seriesVisa: seriesVisa,
                        idVisa: idVisa,
                        dateOfIssueVisa: dateOfIssueVisa,
                        validUntilVisa: validUntilVisa,


                        numOrder: numOrder,
                        orderFrom: orderFrom,
                        orderUntil: orderUntil,

                        typeFunding: typeFunding,
                        numContract: numContract,
                        contractFrom: contractFrom,

                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + "ОВМ ТРОПАРЕВО-НИКУЛИНО" + "/" +
                        "ХОДАТАЙСТВО (ВИЗА) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + "ОВМ ТРОПАРЕВО-НИКУЛИНО" + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ТРОПАРЕВО-НИКУЛИНО.zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ТРОПАРЕВО-НИКУЛИНО.zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                //saveAs(content,nameFile);
            });
    };

//ходатайство ВИЗА ХАМОВНИКИ
    window.generateVisaSolicitaionKhamovnikiTotal = function generate() {
        path = ('../Templates/виза/ходатайство ХАМОВНИКИ.docx')

        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))


                    //dateUntil
                    let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                    // purpose
                    let purpose = ''
                    let purposeS = ''
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purpose = "УЧЕБА"
                            purposeS = "Студент"
                            break
                        case "Краткосрочная учеба":
                            purpose = "КРАТКОСРОЧНАЯ УЧЕБА"
                            purposeS = "Студент"
                            break
                        case "(НТС)":
                            purpose = "НАУЧНО-ТЕХНИЧЕСКИЕ СВЯЗИ (НТС)"
                            purposeS = "НТС"
                            break
                        case "Трудовая деятельность":
                            purpose = "ТРУДОВАЯ ДЕЯТЕЛЬНОСТЬ"
                            purposeS = "Преподаватель"
                            break
                    }

                    // levelEducation
                    let levelEducation = ''
                    switch (document.getElementById('levelEducation' + indexTab).value) {
                        case "Подготовительный факультет (изучаю русский язык)/ The preparatory faculty":
                            levelEducation = 'подготовительный факультет'
                            break
                        case "бакалавриат/bachelor degree":
                            levelEducation = 'бакалавриат'
                            break
                        case "магистратура/master degree":
                            levelEducation = 'магистратура'
                            break
                        case "аспирантура/post-graduate studies":
                            levelEducation = 'аспирантура'
                            break
                    }

                    // course
                    let course = ''
                    switch (document.getElementById('course' + indexTab).value) {
                        case '1':
                            course = ', 1 курс,'
                            break
                        case '2':
                            course = ', 2 курс,'
                            break
                        case '3':
                            course = ', 3 курс,'
                            break
                        case '4':
                            course = ', 4 курс,'
                            break
                        case '5':
                            course = ', 5 курс,'
                            break
                    }

                    // norification
                    let notificationFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('notificationFrom' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('notificationFrom' + indexTab).value).toLocaleDateString() : ''
                    let notificationUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('notificationUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('notificationUntil' + indexTab).value).toLocaleDateString() : ''
                    let issuedBy = document.getElementById('issuedBy' + indexTab).value != '' ? document.getElementById('issuedBy' + indexTab).value : ''

                    // addressResidence
                    let addressResidence = ''
                    switch (document.getElementById('migrationAddress').value) {
                        case "Квартира":
                            addressResidence = document.getElementById('addressResidence' + indexTab).value
                            break
                        default:
                            addressResidence = document.getElementById('migrationAddress').value
                            break
                    }

                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                    // gender
                    let gender = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            gender = 'м.'
                            break
                        case "Женский / Female":
                            gender = 'ж.'
                            break
                    }

                    // registration On
                    let registrationOn = ''
                    switch (document.getElementById('registrationOn').value) {
                        case "Круглов":
                            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                            break
                        case "Морозова":
                            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                            break
                        case "Орлова":
                            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                            break
                    }

                    // visa
                    let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value : ''
                    let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                        ? document.getElementById('idVisa' + indexTab).value : ''

                    let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                    let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                    // order
                    let numOrder = document.getElementById('numOrder' + indexTab).value
                        ? document.getElementById('numOrder' + indexTab).value : ''
                    let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''

                    // faculty
                    let faculty = ''
                    switch (document.getElementById('faculty' + indexTab).value) {
                        case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                            faculty = 'ИИИ:Музфак'
                            break
                        case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                            faculty = 'ИИИ: Худграф'
                            break
                        case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                            faculty = 'ИСГО'
                            break
                        case "Институт филологии / The Institute of Philology":
                            faculty = 'ИФ'
                            break
                        case "Институт иностранных языков / The Institute of Foreign Languages":
                            faculty = 'ИИЯ'
                            break
                        case "Институт международного образования / The Institute of International Education":
                            faculty = 'ИМО'
                            break
                        case "Институт детства / The Institute of Childhood":
                            faculty = 'ИД'
                            break
                        case "Институт биологии и химии / The Institute of Biology and Chemistry":
                            faculty = 'ИБХ'
                            break
                        case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                            faculty = 'ИФТИС'
                            break
                        case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                            faculty = 'ИФКСиЗ'
                            break
                        case "Географический факультет / The Institute of Geography":
                            faculty = 'Геофак'
                            break
                        case "Институт истории и политики / The Institute of History and Politics":
                            faculty = 'ИИП'
                            break
                        case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                            faculty = 'ИМИ'
                            break
                        case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                            faculty = 'Дош.фак.'
                            break
                        case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                            faculty = 'ИПП'
                            break
                        case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                            faculty = 'ИЖКиМ'
                            break
                        case "Институт развития цифрового образования / The Institute of Digital Education Development":
                            faculty = 'ИРЦО'
                            break
                    }


                    // contract
                    let typeFundingDog1 = ""
                    let typeFundingDog2 = ""
                    let typeFundingNap1 = ""
                    let typeFundingNap2 = ""
                    switch (document.getElementById('typeFunding' + indexTab).value) {
                        case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                            typeFundingDog2 = "Договор"
                            typeFundingNap2 = "направление"
                            break
                        case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                            typeFundingDog1 = "Договор"
                            typeFundingNap1 = "направление"
                            break
                    }
                    let numContract = document.getElementById('numContract' + indexTab).value
                        ? document.getElementById('numContract' + indexTab).value : ''
                    let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                        document.getElementById('contractFrom' + indexTab).value : ''



                    doc.setData({
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                        nStud: document.getElementById('nStud' + indexTab).value,
                        grazd: document.getElementById('grazd' + indexTab).value,
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                        dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                        gender: gender,

                        registrationOn: registrationOn,
                        dateUntil: dateUntil,
                        purpose: purpose,
                        purposeS: purposeS,
                        levelEducation: levelEducation,
                        course: course,

                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,

                        seriesVisa: seriesVisa,
                        idVisa: idVisa,
                        dateOfIssueVisa: dateOfIssueVisa,
                        validUntilVisa: validUntilVisa,

                        seriesMigration: document.getElementById('seriesMigration' + indexTab).value,
                        idMigration: document.getElementById('idMigration' + indexTab).value,
                        dateArrivalMigration: new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString(),

                        notificationFrom: notificationFrom,
                        notificationUntil: notificationUntil,
                        issuedBy: issuedBy,

                        addressResidence: addressResidence,
                        faculty: faculty,
                        numOrder: numOrder,
                        orderFrom: orderFrom,
                        typeFundingDog1: typeFundingDog1,
                        typeFundingDog2: typeFundingDog2,
                        typeFundingNap1: typeFundingNap1,
                        typeFundingNap2: typeFundingNap2,

                        numContract: numContract,
                        contractFrom: contractFrom,

                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + "ОВМ ХАМОВНИКИ" + "/" +
                        "ХОДАТАЙСТВО (ВИЗА) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + "ОВМ ХАМОВНИКИ" + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ХАМОВНИКИ.zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ХАМОВНИКИ.zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                //saveAs(content,nameFile);
            });
    };

// выбор функции для ходатайства, относительно выбора ОВМ
    function generateVisaSolicTotal() {
        if (document.getElementById('ovmByRegion').value =='Тропарево-Никулино') {
            generateVisaSolicitaionTroparevoTotal()
        }
        else if (document.getElementById('ovmByRegion').value =='Хамовники') {
            generateVisaSolicitaionKhamovnikiTotal()
        }
    }


//справка
    window.generateVisaReferenceTotal = function generate() {
        path = ('../Templates/виза/справка.docx')
        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))

                    // purpose
                    let purpose = ""
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purpose = "студентом"
                            break
                        case "Краткосрочная учеба":
                            purpose = "студентом"
                            break
                        case "(НТС)":
                            purpose = "приглашенным гостем (НТС)"
                            break
                        case "Трудовая деятельность":
                            purpose = "преподавателем"
                            break
                    }

                    // faculty
                    let faculty = ''
                    switch (document.getElementById('faculty' + indexTab).value) {
                        case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                            faculty = 'Института изящных искусств: Факультета музыкального искусства'
                            break
                        case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                            faculty = 'Института изящных искусств: Художественно-графического факультета'
                            break
                        case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                            faculty = 'Института социально-гуманитарного образования'
                            break
                        case "Институт филологии / The Institute of Philology":
                            faculty = 'Института филологии'
                            break
                        case "Институт иностранных языков / The Institute of Foreign Languages":
                            faculty = 'Института иностранных языков'
                            break
                        case "Институт международного образования / The Institute of International Education":
                            faculty = 'Института международного образования'
                            break
                        case "Институт детства / The Institute of Childhood":
                            faculty = 'Института детства'
                            break
                        case "Институт биологии и химии / The Institute of Biology and Chemistry":
                            faculty = 'Института биологии и химии'
                            break
                        case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                            faculty = 'Института физики, технологии и информационных систем'
                            break
                        case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                            faculty = 'Института физической культуры, спорта и здоровья'
                            break
                        case "Географический факультет / The Institute of Geography":
                            faculty = 'Географического факультета'
                            break
                        case "Институт истории и политики / The Institute of History and Politics":
                            faculty = 'Института истории и политики'
                            break
                        case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                            faculty = 'Института математики и информатики'
                            break
                        case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                            faculty = 'Факультета дошкольной педагогики и психологии'
                            break
                        case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                            faculty = 'Института педагогики и психологии'
                            break
                        case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                            faculty = 'Института журналистики, коммуникаций и медиаобразования'
                            break
                        case "Институт развития цифрового образования / The Institute of Digital Education Development":
                            faculty = 'Института развития цифрового образования'
                            break

                    }


                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''



                    // OVM
                    let ovmByRegion = ''
                    switch (document.getElementById('ovmByRegion').value) {
                        case "Тропарево-Никулино":
                            ovmByRegion = 'ОМВД России по району Тропарево-Никулино г.Москвы'
                            break
                        case "Хамовники":
                            ovmByRegion = 'ОМВД России по району Хамовники г.Москвы'
                            break
                    }

                    // registration On
                    let registrationOn = ''
                    switch (document.getElementById('registrationOn').value) {
                        case "Круглов":
                            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                            break
                        case "Морозова":
                            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                            break
                        case "Орлова":
                            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                            break
                    }

                    let dateUnt =  document.getElementById('dateUntil' + indexTab).value ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''


                    doc.setData({

                        grazd: document.getElementById('grazd' + indexTab).value.toUpperCase(),
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value.toUpperCase(),
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value.toUpperCase(),
                        purpose: purpose,
                        faculty: faculty.toUpperCase(),
                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,
                        dateUntil: dateUnt,
                        ovmByRegion: ovmByRegion,
                        registrationOn: registrationOn,
                        nStud: document.getElementById('nStud' + indexTab).value,
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + "/" +

                        "СПРАВКА - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" СПРАВКА - "  + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" СПРАВКА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                // saveAs(content,nameFile);
            });
    };

//опись ВИЗА
    window.generateInventoryVisaTotal = function generate() {
        path = ('../Templates/виза/опись виза.docx')

        let students = []

        // registration On
        let registrationOn = ''
        switch (document.getElementById('registrationOn').value) {
            case "Круглов":
                registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                break
            case "Морозова":
                registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                break
            case "Орлова":
                registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                break
        }

        // ovmByRegion
        let ovmByRegion = ''
        switch (document.getElementById('ovmByRegion').value) {
            case 'Тропарево-Никулино':
                ovmByRegion = 'ОМВД России по району Тропарево-Никулино г. Москвы'
                break
            case 'Хамовники':
                ovmByRegion = 'ОМВД России по району Хамовники г. Москвы'
                break
            case 'Алексеевский':
                ovmByRegion = 'ОМВД России по Алексеевскому району г.Москвы'
                break
            case 'Войковский':
                ovmByRegion = 'ОВМ МВД России по району Войковский г. Москвы'
                break
            case 'МУ МВД РФ Люберецкое':
                ovmByRegion = 'ОВМ МУ МВД РФ "Люберецкое"'
                break
        }

        // nStud
        let nStud1 = document.getElementById('nStud1').value
        let tir = ''
        let nStud2 = ''
        if (countTab()>1) {
            nStud2 = document.getElementById('nStud' + (lastTab()-1)).value
            tir = '-'
        }

        for (let i =0; i<countTab();i++) {
            let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
            let elem = tabs[i]
            let indexTab = parseInt(elem.id.match(/\d+/))



            let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''

            let dateUntil = document.getElementById('dateUntil' + indexTab).value != '-' ?
                new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

            //students
            students.push({
                nStud: document.getElementById('nStud' + indexTab).value,
                lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                firstNameRu: document.getElementById('firstNameRu'+indexTab).value.toUpperCase(),
                patronymicRu: document.getElementById('patronymicRu'+indexTab).value.toUpperCase(),
                dateOfIssueVisa: dateOfIssueVisa,
                dateUntil: dateUntil,
                grazd: document.getElementById('grazd' + indexTab).value,
                phone: document.getElementById('phone' + indexTab).value,
                mail: document.getElementById('mail' + indexTab).value,
            })
        }


        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                doc.render({
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    nStud1: nStud1,
                    nStud2: nStud2,
                    tir: tir,
                    ovmByRegion: ovmByRegion,
                    'students': students,
                    registrationOn: registrationOn,
                })


                var out = doc.getZip().generate({
                });



                let nameFile = ''
                if (countTab()==1) {
                    nameFile = "ОПИСЬ (ВИЗА) - студент " +
                        document.getElementById('nStud1').value
                        +" - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                        +" - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                }
                else {
                    nameFile = "ОПИСЬ (ВИЗА) - студенты " +
                        document.getElementById('nStud1').value
                        +"-"+
                        document.getElementById('nStud'+(lastTab()-1)).value
                        +" - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                        +" - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                }



                zipTotal.file(nameFile, out, {base64: true})

                // Output the document using Data-URI
                //saveAs(out, nameFile);
                var content = zipTotal.generate({type: 'blob'})

                let nameZip = ''
                if (countTab()==1) {
                    nameZip = document.getElementById('nStud1').value
                        +" Студент.zip"
                }
                else {
                    nameZip = document.getElementById('nStud1').value + '-'+
                        document.getElementById('nStud'+(lastTab()-1)).value
                        +" Студенты.zip"
                }

                saveAs(content, nameZip)
            }
        );
    };



    setTimeout(generateVisaApplicationTotal, 30)
    setTimeout(generateVisaSolicTotal, 30)
    setTimeout(generateVisaReferenceTotal, 30)
    setTimeout(generateInventoryVisaTotal, 300)

}





// РЕГИСТРАЦИЯ + ВИЗА
function generateRegVisa() {
    var zipTotal = new PizZip();

    //уведомление РЕГИСТРАЦИЯ
    window.generateRegNotifTot = function generate() {
        let ovmRg = document.getElementById('ovmByRegion').value
        let rgOn = document.getElementById('registrationOn').value
        path = (`../Templates/регистрация/уведомление ${ovmRg} ${rgOn}.docx`)

        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))

                    // lastName {lN1-27}
                    let lN = document.getElementById('lastNameRu'+indexTab).value.toUpperCase().split('')
                    let lN1 = (lN[0]) ? (lN[0]) : ''
                    let lN2 = (lN[1]) ? (lN[1]) : ''
                    let lN3 = (lN[2]) ? (lN[2]) : ''
                    let lN4 = (lN[3]) ? (lN[3]) : ''
                    let lN5 = (lN[4]) ? (lN[4]) : ''
                    let lN6 = (lN[5]) ? (lN[5]) : ''
                    let lN7 = (lN[6]) ? (lN[6]) : ''
                    let lN8 = (lN[7]) ? (lN[7]) : ''
                    let lN9 = (lN[8]) ? (lN[8]) : ''
                    let lN10 = (lN[9]) ? (lN[9]) : ''
                    let lN11 = (lN[10]) ? (lN[10]) : ''
                    let lN12 = (lN[11]) ? (lN[11]) : ''
                    let lN13 = (lN[12]) ? (lN[12]) : ''
                    let lN14 = (lN[13]) ? (lN[13]) : ''
                    let lN15 = (lN[14]) ? (lN[14]) : ''
                    let lN16 = (lN[15]) ? (lN[15]) : ''
                    let lN17 = (lN[16]) ? (lN[16]) : ''
                    let lN18 = (lN[17]) ? (lN[17]) : ''
                    let lN19 = (lN[18]) ? (lN[18]) : ''
                    let lN20 = (lN[19]) ? (lN[19]) : ''
                    let lN21 = (lN[20]) ? (lN[20]) : ''
                    let lN22 = (lN[21]) ? (lN[21]) : ''
                    let lN23 = (lN[22]) ? (lN[22]) : ''
                    let lN24 = (lN[23]) ? (lN[23]) : ''
                    let lN25 = (lN[24]) ? (lN[24]) : ''
                    let lN26 = (lN[25]) ? (lN[25]) : ''
                    let lN27 = (lN[26]) ? (lN[26]) : ''

                    // firstName {fN1-f27}
                    let fN = document.getElementById('firstNameRu' + indexTab).value.toUpperCase().split('')
                    let fN1 = (fN[0]) ? (fN[0]) : ''
                    let fN2 = (fN[1]) ? (fN[1]) : ''
                    let fN3 = (fN[2]) ? (fN[2]) : ''
                    let fN4 = (fN[3]) ? (fN[3]) : ''
                    let fN5 = (fN[4]) ? (fN[4]) : ''
                    let fN6 = (fN[5]) ? (fN[5]) : ''
                    let fN7 = (fN[6]) ? (fN[6]) : ''
                    let fN8 = (fN[7]) ? (fN[7]) : ''
                    let fN9 = (fN[8]) ? (fN[8]) : ''
                    let fN10 = (fN[9]) ? (fN[9]) : ''
                    let fN11 = (fN[10]) ? (fN[10]) : ''
                    let fN12 = (fN[11]) ? (fN[11]) : ''
                    let fN13 = (fN[12]) ? (fN[12]) : ''
                    let fN14 = (fN[13]) ? (fN[13]) : ''
                    let fN15 = (fN[14]) ? (fN[14]) : ''
                    let fN16 = (fN[15]) ? (fN[15]) : ''
                    let fN17 = (fN[16]) ? (fN[16]) : ''
                    let fN18 = (fN[17]) ? (fN[17]) : ''
                    let fN19 = (fN[18]) ? (fN[18]) : ''
                    let fN20 = (fN[19]) ? (fN[19]) : ''
                    let fN21 = (fN[20]) ? (fN[20]) : ''
                    let fN22 = (fN[21]) ? (fN[21]) : ''
                    let fN23 = (fN[22]) ? (fN[22]) : ''
                    let fN24 = (fN[23]) ? (fN[23]) : ''
                    let fN25 = (fN[24]) ? (fN[24]) : ''
                    let fN26 = (fN[25]) ? (fN[25]) : ''
                    let fN27 = (fN[26]) ? (fN[26]) : ''

                    // patronymic {patr1-24}
                    let patr = document.getElementById('patronymicRu' + indexTab).value.toUpperCase().split('')
                    let patr1 = (patr[0]) ? (patr[0]) : ''
                    let patr2 = (patr[1]) ? (patr[1]) : ''
                    let patr3 = (patr[2]) ? (patr[2]) : ''
                    let patr4 = (patr[3]) ? (patr[3]) : ''
                    let patr5 = (patr[4]) ? (patr[4]) : ''
                    let patr6 = (patr[5]) ? (patr[5]) : ''
                    let patr7 = (patr[6]) ? (patr[6]) : ''
                    let patr8 = (patr[7]) ? (patr[7]) : ''
                    let patr9 = (patr[8]) ? (patr[8]) : ''
                    let patr10 = (patr[9]) ? (patr[9]) : ''
                    let patr11 = (patr[10]) ? (patr[10]) : ''
                    let patr12 = (patr[11]) ? (patr[11]) : ''
                    let patr13 = (patr[12]) ? (patr[12]) : ''
                    let patr14 = (patr[13]) ? (patr[13]) : ''
                    let patr15 = (patr[14]) ? (patr[14]) : ''
                    let patr16 = (patr[15]) ? (patr[15]) : ''
                    let patr17 = (patr[16]) ? (patr[16]) : ''
                    let patr18 = (patr[17]) ? (patr[17]) : ''
                    let patr19 = (patr[18]) ? (patr[18]) : ''
                    let patr20 = (patr[19]) ? (patr[19]) : ''
                    let patr21 = (patr[20]) ? (patr[20]) : ''
                    let patr22 = (patr[21]) ? (patr[21]) : ''
                    let patr23 = (patr[22]) ? (patr[22]) : ''
                    let patr24 = (patr[23]) ? (patr[23]) : ''

                    // grazd {grazd1-25}
                    let grazd = document.getElementById('grazd' + indexTab).value.toUpperCase().split('')
                    let grazd1 = (grazd[0]) ? (grazd[0]) : ''
                    let grazd2 = (grazd[1]) ? (grazd[1]) : ''
                    let grazd3 = (grazd[2]) ? (grazd[2]) : ''
                    let grazd4 = (grazd[3]) ? (grazd[3]) : ''
                    let grazd5 = (grazd[4]) ? (grazd[4]) : ''
                    let grazd6 = (grazd[5]) ? (grazd[5]) : ''
                    let grazd7 = (grazd[6]) ? (grazd[6]) : ''
                    let grazd8 = (grazd[7]) ? (grazd[7]) : ''
                    let grazd9 = (grazd[8]) ? (grazd[8]) : ''
                    let grazd10 = (grazd[9]) ? (grazd[9]) : ''
                    let grazd11 = (grazd[10]) ? (grazd[10]) : ''
                    let grazd12 = (grazd[11]) ? (grazd[11]) : ''
                    let grazd13 = (grazd[12]) ? (grazd[12]) : ''
                    let grazd14 = (grazd[13]) ? (grazd[13]) : ''
                    let grazd15 = (grazd[14]) ? (grazd[14]) : ''
                    let grazd16 = (grazd[15]) ? (grazd[15]) : ''
                    let grazd17 = (grazd[16]) ? (grazd[16]) : ''
                    let grazd18 = (grazd[17]) ? (grazd[17]) : ''
                    let grazd19 = (grazd[18]) ? (grazd[18]) : ''
                    let grazd20 = (grazd[19]) ? (grazd[19]) : ''
                    let grazd21 = (grazd[20]) ? (grazd[20]) : ''
                    let grazd22 = (grazd[21]) ? (grazd[21]) : ''
                    let grazd23 = (grazd[22]) ? (grazd[22]) : ''
                    let grazd24 = (grazd[23]) ? (grazd[23]) : ''
                    let grazd25 = (grazd[24]) ? (grazd[24]) : ''

                    // dateOfBirth {dOB1-8}
                    let dOB = new Date(document.getElementById('dateOfBirth'+indexTab).value).toLocaleDateString().split('.')
                    let dOB1 = (dOB) ? (dOB[0][0]) : ''
                    let dOB2 = (dOB) ? (dOB[0][1]) : ''
                    let dOB3 = (dOB) ? (dOB[1][0]) : ''
                    let dOB4 = (dOB) ? (dOB[1][1]) : ''
                    let dOB5 = (dOB) ? (dOB[2][0]) : ''
                    let dOB6 = (dOB) ? (dOB[2][1]) : ''
                    let dOB7 = (dOB) ? (dOB[2][2]) : ''
                    let dOB8 = (dOB) ? (dOB[2][3]) : ''

                    // gender
                    let genM = ''
                    let genW = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            genM = 'X'
                            break
                        case "Женский / Female":
                            genW = 'X'
                            break
                    }

                    // placeStateBirth {pSBS1-24} для страны И {pSBG1-24} для города
                    let pSBS = document.getElementById('placeStateBirth'+indexTab).value.toUpperCase().split(', ')
                    let pSBS1 = (pSBS[0][0]) ? (pSBS[0][0]) : ''
                    let pSBS2 = (pSBS[0][1]) ? (pSBS[0][1]) : ''
                    let pSBS3 = (pSBS[0][2]) ? (pSBS[0][2]) : ''
                    let pSBS4 = (pSBS[0][3]) ? (pSBS[0][3]) : ''
                    let pSBS5 = (pSBS[0][4]) ? (pSBS[0][4]) : ''
                    let pSBS6 = (pSBS[0][5]) ? (pSBS[0][5]) : ''
                    let pSBS7 = (pSBS[0][6]) ? (pSBS[0][6]) : ''
                    let pSBS8 = (pSBS[0][7]) ? (pSBS[0][7]) : ''
                    let pSBS9 = (pSBS[0][8]) ? (pSBS[0][8]) : ''
                    let pSBS10 = (pSBS[0][9]) ? (pSBS[0][9]) : ''
                    let pSBS11 = (pSBS[0][10]) ? (pSBS[0][10]) : ''
                    let pSBS12 = (pSBS[0][11]) ? (pSBS[0][11]) : ''
                    let pSBS13 = (pSBS[0][12]) ? (pSBS[0][12]) : ''
                    let pSBS14 = (pSBS[0][13]) ? (pSBS[0][13]) : ''
                    let pSBS15 = (pSBS[0][14]) ? (pSBS[0][14]) : ''
                    let pSBS16 = (pSBS[0][15]) ? (pSBS[0][15]) : ''
                    let pSBS17 = (pSBS[0][16]) ? (pSBS[0][16]) : ''
                    let pSBS18 = (pSBS[0][17]) ? (pSBS[0][17]) : ''
                    let pSBS19 = (pSBS[0][18]) ? (pSBS[0][18]) : ''
                    let pSBS20 = (pSBS[0][19]) ? (pSBS[0][19]) : ''
                    let pSBS21 = (pSBS[0][20]) ? (pSBS[0][20]) : ''
                    let pSBS22 = (pSBS[0][21]) ? (pSBS[0][21]) : ''
                    let pSBS23 = (pSBS[0][22]) ? (pSBS[0][22]) : ''
                    let pSBS24 = (pSBS[0][23]) ? (pSBS[0][23]) : ''
                    let pSBG1 = ""
                    let pSBG2 = ""
                    let pSBG3 = ""
                    let pSBG4 = ""
                    let pSBG5 = ""
                    let pSBG6 = ""
                    let pSBG7 = ""
                    let pSBG8 = ""
                    let pSBG9 = ""
                    let pSBG10 = ""
                    let pSBG11 = ""
                    let pSBG12 = ""
                    let pSBG13 = ""
                    let pSBG14 = ""
                    let pSBG15 = ""
                    let pSBG16 = ""
                    let pSBG17 = ""
                    let pSBG18 = ""
                    let pSBG19 = ""
                    let pSBG20 = ""
                    let pSBG21 = ""
                    let pSBG22 = ""
                    let pSBG23 = ""
                    let pSBG24 = ""
                    if (pSBS.length>1) {
                        pSBG1 = (pSBS[1][0]) ? (pSBS[1][0]) : ''
                        pSBG2 = (pSBS[1][1]) ? (pSBS[1][1]) : ''
                        pSBG3 = (pSBS[1][2]) ? (pSBS[1][2]) : ''
                        pSBG4 = (pSBS[1][3]) ? (pSBS[1][3]) : ''
                        pSBG5 = (pSBS[1][4]) ? (pSBS[1][4]) : ''
                        pSBG6 = (pSBS[1][5]) ? (pSBS[1][5]) : ''
                        pSBG7 = (pSBS[1][6]) ? (pSBS[1][6]) : ''
                        pSBG8 = (pSBS[1][7]) ? (pSBS[1][7]) : ''
                        pSBG9 = (pSBS[1][8]) ? (pSBS[1][8]) : ''
                        pSBG10 = (pSBS[1][9]) ? (pSBS[1][9]) : ''
                        pSBG11 = (pSBS[1][10]) ? (pSBS[1][10]) : ''
                        pSBG12 = (pSBS[1][11]) ? (pSBS[1][11]) : ''
                        pSBG13 = (pSBS[1][12]) ? (pSBS[1][12]) : ''
                        pSBG14 = (pSBS[1][13]) ? (pSBS[1][13]) : ''
                        pSBG15 = (pSBS[1][14]) ? (pSBS[1][14]) : ''
                        pSBG16 = (pSBS[1][15]) ? (pSBS[1][15]) : ''
                        pSBG17 = (pSBS[1][16]) ? (pSBS[1][16]) : ''
                        pSBG18 = (pSBS[1][17]) ? (pSBS[1][17]) : ''
                        pSBG19 = (pSBS[1][18]) ? (pSBS[1][18]) : ''
                        pSBG20 = (pSBS[1][19]) ? (pSBS[1][19]) : ''
                        pSBG21 = (pSBS[1][20]) ? (pSBS[1][20]) : ''
                        pSBG22 = (pSBS[1][21]) ? (pSBS[1][21]) : ''
                        pSBG23 = (pSBS[1][22]) ? (pSBS[1][22]) : ''
                        pSBG24 = (pSBS[1][23]) ? (pSBS[1][23]) : ''
                    }

                    // series {sP1-4}
                    let sP = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value.toUpperCase() : ''
                    let sP1 = ''
                    let sP2 = ''
                    let sP3 = ''
                    let sP4 = ''
                    if (sP) {
                        switch (sP.length) {
                            case 1:
                                sP4 = sP[0]
                                break
                            case 2:
                                sP3 = sP[0]
                                sP4 = sP[1]
                                break
                            case 3:
                                sP2 = sP[0]
                                sP3 = sP[1]
                                sP4 = sP[2]
                                break
                            case 4:
                                sP1 = sP[0]
                                sP2 = sP[1]
                                sP3 = sP[2]
                                sP4 = sP[3]
                        }
                    }


                    // idPassport {iP1-10}
                    let idP = document.getElementById('idPassport'+indexTab).value.toUpperCase()
                    let idP1 = (idP[0]) ? (idP[0]) : ''
                    let idP2 = (idP[1]) ? (idP[1]) : ''
                    let idP3 = (idP[2]) ? (idP[2]) : ''
                    let idP4 = (idP[3]) ? (idP[3]) : ''
                    let idP5 = (idP[4]) ? (idP[4]) : ''
                    let idP6 = (idP[5]) ? (idP[5]) : ''
                    let idP7 = (idP[6]) ? (idP[6]) : ''
                    let idP8 = (idP[7]) ? (idP[7]) : ''
                    let idP9 = (idP[8]) ? (idP[8]) : ''
                    let idP10 = (idP[9]) ? (idP[9]) : ''

                    // dateOfIssue {dOI1-8}
                    let dOI = new Date(document.getElementById('dateOfIssue'+indexTab).value).toLocaleDateString().split('.')
                    let dOI1 = (dOI) ? (dOI[0][0]) : ''
                    let dOI2 = (dOI) ? (dOI[0][1]) : ''
                    let dOI3 = (dOI) ? (dOI[1][0]) : ''
                    let dOI4 = (dOI) ? (dOI[1][1]) : ''
                    let dOI5 = (dOI) ? (dOI[2][0]) : ''
                    let dOI6 = (dOI) ? (dOI[2][1]) : ''
                    let dOI7 = (dOI) ? (dOI[2][2]) : ''
                    let dOI8 = (dOI) ? (dOI[2][3]) : ''


                    // validUntil {vU1-8}
                    let vU = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString().split(".") : ''
                    let vU1 = (vU) ? (vU[0][0]) : ''
                    let vU2 = (vU) ? (vU[0][1]) : ''
                    let vU3 = (vU) ? (vU[1][0]) : ''
                    let vU4 = (vU) ? (vU[1][1]) : ''
                    let vU5 = (vU) ? (vU[2][0]) : ''
                    let vU6 = (vU) ? (vU[2][1]) : ''
                    let vU7 = (vU) ? (vU[2][2]) : ''
                    let vU8 = (vU) ? (vU[2][3]) : ''


                    // typeVisa
                    let tVV = ''
                    let tVJ = ''
                    let tVP = ''
                    switch (document.getElementById('typeVisa' + indexTab).value) {
                        case "ВИЗА":
                            tVV = 'X'
                            break
                        case "(ВНЖ) ВИД НА ЖИТЕЛЬСТВО РФ":
                            tVJ = 'X'
                            break
                        case "(РВП) РАЗРЕШЕНИЕ НА ВРЕМЕННОЕ ПРОЖИВАНИЕ РФ":
                            tVP = 'X'
                            break
                    }

                    // seriesVisa {sV1-4}
                    let sV = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value.toUpperCase() : ''
                    let sV1 = ''
                    let sV2 = ''
                    let sV3 = ''
                    let sV4 = ''
                    if (sV) {
                        switch (sV.length) {
                            case 1:
                                sV4 = sV[0]
                                break
                            case 2:
                                sV3 = sV[0]
                                sV4 = sV[1]
                                break
                            case 3:
                                sV2 = sV[0]
                                sV3 = sV[1]
                                sV4 = sV[2]
                                break
                            case 4:
                                sV1 = sV[0]
                                sV2 = sV[1]
                                sV3 = sV[2]
                                sV4 = sV[3]
                        }
                    }

                    // idVisa {idV1-15}
                    let idV = document.getElementById('idVisa' + indexTab).value ? document.getElementById('idVisa' + indexTab).value.toUpperCase() : ''
                    let idV1 = (idV[0]) ? (idV[0]) : ''
                    let idV2 = (idV[1]) ? (idV[1]) : ''
                    let idV3 = (idV[2]) ? (idV[2]) : ''
                    let idV4 = (idV[3]) ? (idV[3]) : ''
                    let idV5 = (idV[4]) ? (idV[4]) : ''
                    let idV6 = (idV[5]) ? (idV[5]) : ''
                    let idV7 = (idV[6]) ? (idV[6]) : ''
                    let idV8 = (idV[7]) ? (idV[7]) : ''
                    let idV9 = (idV[8]) ? (idV[8]) : ''
                    let idV10 = (idV[9]) ? (idV[9]) : ''
                    let idV11 = (idV[10]) ? (idV[10]) : ''
                    let idV12 = (idV[11]) ? (idV[11]) : ''
                    let idV13 = (idV[12]) ? (idV[12]) : ''
                    let idV14 = (idV[13]) ? (idV[13]) : ''
                    let idV15 = (idV[14]) ? (idV[14]) : ''

                    // dateOfIssueVisa {dOIV1-8}
                    let dOIV = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString().split(".") : ''
                    let dOIV1 = (dOIV) ? (dOIV[0][0]) : ''
                    let dOIV2 = (dOIV) ? (dOIV[0][1]) : ''
                    let dOIV3 = (dOIV) ? (dOIV[1][0]) : ''
                    let dOIV4 = (dOIV) ? (dOIV[1][1]) : ''
                    let dOIV5 = (dOIV) ? (dOIV[2][0]) : ''
                    let dOIV6 = (dOIV) ? (dOIV[2][1]) : ''
                    let dOIV7 = (dOIV) ? (dOIV[2][2]) : ''
                    let dOIV8 = (dOIV) ? (dOIV[2][3]) : ''

                    // validUntilVisa {vUV1-8}
                    let vUV = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString().split(".") : ''
                    let vUV1 = (vUV) ? (vUV[0][0]) : ''
                    let vUV2 = (vUV) ? (vUV[0][1]) : ''
                    let vUV3 = (vUV) ? (vUV[1][0]) : ''
                    let vUV4 = (vUV) ? (vUV[1][1]) : ''
                    let vUV5 = (vUV) ? (vUV[2][0]) : ''
                    let vUV6 = (vUV) ? (vUV[2][1]) : ''
                    let vUV7 = (vUV) ? (vUV[2][2]) : ''
                    let vUV8 = (vUV) ? (vUV[2][3]) : ''

                    // purpose
                    let purposeG = ""
                    let purposeR = ""
                    let purposeU = ""
                    let purp1 = ''
                    let purp2 = ''
                    let purp3 = ''
                    let purp4 = ''
                    let purp5 = ''
                    let purp6 = ''
                    let purp7 = ''
                    let purp8 = ''
                    let purp9 = ''
                    let purp10 = ''
                    let purp11 = ''
                    let purp12 = ''
                    let purp13 = ''
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purposeU = "X"
                            purp1 = 'С'; purp2='Т'; purp3='У'; purp4="Д"; purp5 ='Е';purp6 = 'Н';purp7='Т'
                            break
                        case "Краткосрочная учеба":
                            purposeU = "X"
                            purp1 = 'С'; purp2='Т'; purp3='У'; purp4="Д"; purp5 ='Е';purp6 = 'Н';purp7='Т'
                            break
                        case "(НТС)":
                            purposeG = "X"
                            purp1 = 'Н'; purp2='Т'; purp3='С';
                            break
                        case "Трудовая деятельность":
                            purposeR = "X"
                            purp1='П';purp2='Р';purp3='Е';purp4="П";purp5='О';purp6='Д';purp7='А';purp8='В';purp9='А';purp10='Т';purp11='Е';purp12='Л';purp13='Ь';
                            break
                    }

                    // dateArrivalMigration {dAM1-8}
                    let dAM = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString().split(".") : ''
                    let dAM1 = (dAM) ? (dAM[0][0]) : ''
                    let dAM2 = (dAM) ? (dAM[0][1]) : ''
                    let dAM3 = (dAM) ? (dAM[1][0]) : ''
                    let dAM4 = (dAM) ? (dAM[1][1]) : ''
                    let dAM5 = (dAM) ? (dAM[2][0]) : ''
                    let dAM6 = (dAM) ? (dAM[2][1]) : ''
                    let dAM7 = (dAM) ? (dAM[2][2]) : ''
                    let dAM8 = (dAM) ? (dAM[2][3]) : ''

                    // dateUntil {dU1-8}
                    let dU = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString().split(".") : ''
                    let dU1 = (dU) ? (dU[0][0]) : ''
                    let dU2 = (dU) ? (dU[0][1]) : ''
                    let dU3 = (dU) ? (dU[1][0]) : ''
                    let dU4 = (dU) ? (dU[1][1]) : ''
                    let dU5 = (dU) ? (dU[2][0]) : ''
                    let dU6 = (dU) ? (dU[2][1]) : ''
                    let dU7 = (dU) ? (dU[2][2]) : ''
                    let dU8 = (dU) ? (dU[2][3]) : ''

                    // seriesMigration {sM1-4}
                    let sM = document.getElementById('seriesMigration' + indexTab).value
                        ? document.getElementById('seriesMigration' + indexTab).value.toUpperCase() : ''
                    let sM1 = ''
                    let sM2 = ''
                    let sM3 = ''
                    let sM4 = ''
                    if (sM) {
                        switch (sM.length) {
                            case 1:
                                sM4 = sM[0]
                                break
                            case 2:
                                sM3 = sM[0]
                                sM4 = sM[1]
                                break
                            case 3:
                                sM2 = sM[0]
                                sM3 = sM[1]
                                sM4 = sM[2]
                                break
                            case 4:
                                sM1 = sM[0]
                                sM2 = sM[1]
                                sM3 = sM[2]
                                sM4 = sM[3]
                        }
                    }

                    //idMigration {iM1-11}
                    let iM = document.getElementById('idMigration' + indexTab).value
                        ? document.getElementById('idMigration' + indexTab).value.toUpperCase() : ''
                    let iM1 = (iM[0]) ? (iM[0]) : ''
                    let iM2 = (iM[1]) ? (iM[1]) : ''
                    let iM3 = (iM[2]) ? (iM[2]) : ''
                    let iM4 = (iM[3]) ? (iM[3]) : ''
                    let iM5 = (iM[4]) ? (iM[4]) : ''
                    let iM6 = (iM[5]) ? (iM[5]) : ''
                    let iM7 = (iM[6]) ? (iM[6]) : ''
                    let iM8 = (iM[7]) ? (iM[7]) : ''
                    let iM9 = (iM[8]) ? (iM[8]) : ''
                    let iM10 = (iM[9]) ? (iM[9]) : ''
                    let iM11 = (iM[10]) ? (iM[10]) : ''

                    // addressHostel {aHG1-18} {aHU1-20} {aHD1-15} {aHK1-5}
                    let aHG1 = ''
                    let aHG2 = ''
                    let aHG3 = ''
                    let aHG4 = ''
                    let aHG5 = ''
                    let aHG6 = ''
                    let aHG7 = ''
                    let aHG8 = ''
                    let aHG9 = ''
                    let aHG10 = ''
                    let aHG11 = ''
                    let aHG12 = ''
                    let aHG13 = ''
                    let aHG14 = ''
                    let aHG15 = ''
                    let aHG16 = ''
                    let aHG17 = ''
                    let aHG18 = ''
                    let aHU1 = ''
                    let aHU2 = ''
                    let aHU3 = ''
                    let aHU4 = ''
                    let aHU5 = ''
                    let aHU6 = ''
                    let aHU7 = ''
                    let aHU8 = ''
                    let aHU9 = ''
                    let aHU10 = ''
                    let aHU11 = ''
                    let aHU12 = ''
                    let aHU13 = ''
                    let aHU14 = ''
                    let aHU15 = ''
                    let aHU16 = ''
                    let aHU17 = ''
                    let aHU18 = ''
                    let aHU19 = ''
                    let aHU20 = ''
                    let aHD1 = ""
                    let aHD2 = ""
                    let aHD3 = ""
                    let aHD4 = ""
                    let aHD5 = ""
                    let aHD6 = ""
                    let aHD7 = ""
                    let aHD8 = ""
                    let aHD9 = ""
                    let aHD10 = ""
                    let aHD11 = ""
                    let aHD12 = ""
                    let aHD13 = ""
                    let aHD14 = ""
                    let aHD15 = ""
                    let aHK1 = ''
                    let aHK2 = ''
                    let aHK3 = ''
                    let aHK4 = ''
                    let aHK5 = ''

                    switch (document.getElementById('addressHostel'+indexTab).value) {
                        case "г. Москва, проспект Вернадского, 88 к. 1 (ОБЩЕЖИТИЕ №1)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '1';
                            break
                        case "г. Москва, проспект Вернадского, 88 к. 2 (ОБЩЕЖИТИЕ №2)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '2';
                            break
                        case "г. Москва, проспект Вернадского, 88 к. 3 (ОБЩЕЖИТИЕ №3)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'П'; aHU2 = 'Р'; aHU3 = 'О'; aHU4 = 'С'; aHU5 = 'П'; aHU6 = 'Е'; aHU7 = 'К'; aHU8 = 'Т'; aHU10 = 'В'; aHU11 = 'Е'; aHU12 = 'Р'; aHU13 = 'Н'; aHU14 = 'А'; aHU15 = 'Д'; aHU16 = 'С'; aHU17 = 'К'; aHU18 = 'О'; aHU19 = 'Г'; aHU20 = 'О';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '8'; aHD6 = '8'; aHD7 = ''; aHD8 = 'К'; aHD9 = 'О'; aHD10 = 'Р'; aHD11 = 'П'; aHD12 = 'У'; aHD13 = 'С'; aHD14 = ''; aHD15 = '3';
                            break
                        case "г. Москва, улица Космонавтов, д. 13 (ОБЩЕЖИТИЕ №4)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'К'; aHU2 = 'О'; aHU3 = 'С'; aHU4 = 'М'; aHU5 = 'О'; aHU6 = 'Н'; aHU7 = 'А'; aHU8 = 'В'; aHU9 = 'Т'; aHU10 = 'О'; aHU11 = 'В';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '1'; aHD6 = '3';
                            break
                        case "г. Москва, улица Космонавтов, д. 9 (ОБЩЕЖИТИЕ №5)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'К'; aHU2 = 'О'; aHU3 = 'С'; aHU4 = 'М'; aHU5 = 'О'; aHU6 = 'Н'; aHU7 = 'А'; aHU8 = 'В'; aHU9 = 'Т'; aHU10 = 'О'; aHU11 = 'В';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '9';
                            break
                        case "г. Москва, улица Клары Цеткин, д. 25 (ОБЩЕЖИТИЕ №6)":
                            aHG1 = 'М'; aHG2 = "О"; aHG3 = "С"; aHG4 = "К"; aHG5 = "В"; aHG6 = "А"
                            aHU1 = 'К'; aHU2 = 'Л'; aHU3 = 'А'; aHU4 = 'Р'; aHU5 = 'Ы'; aHU6 = ''; aHU7 = 'Ц'; aHU8 = 'Е'; aHU9 = 'Т'; aHU10 = 'К'; aHU11 = 'И'; aHU12 = 'Н';
                            aHD1 = 'Д'; aHD2 = 'О'; aHD3 = 'М'; aHD4 = ''; aHD5 = '2'; aHD6 = '5';
                            break
                        case "Московская область, г. Люберцы, ул. Мира, д.7 (ОБЩЕЖИТИЕ №7)":
                            aHG1 = 'М'; aHG2 = 'О'; aHG3 = 'С'; aHG4 = 'К'; aHG5 = 'О'; aHG6 = 'В'; aHG7 = 'С'; aHG8 = 'К'; aHG9 = 'А'; aHG10 = 'Я'; aHG12 = 'О'; aHG13 = 'Б'; aHG14 = 'Л'; aHG15 = 'А'; aHG16 = 'С'; aHG17 = 'Т'; aHG18 = 'Ь';
                            aHU1 = 'Л'; aHU2 = 'Ю'; aHU3 = 'Б'; aHU4 = 'Е'; aHU5 = 'Р'; aHU6 = 'Ц'; aHU7 = 'Ы';
                            aHD1 = 'М'; aHD2 = 'И'; aHD3 = 'Р'; aHD4 = 'А';
                            aHK1 = 'Д'; aHK2 = 'О'; aHK3 = 'М'; aHK4 = ''; aHK5 = '7';
                            break
                    }

                    // migrationAddress {mAO1-18} {mAG1-7} {mAU1-20} {mAD1-2} {mAK1}
                    let mAO1 = ""
                    let mAO2 = ""
                    let mAO3 = ""
                    let mAO4 = ""
                    let mAO5 = ""
                    let mAO6 = ""
                    let mAO7 = ""
                    let mAO8 = ""
                    let mAO9 = ""
                    let mAO10 = ""
                    let mAO11 = ""
                    let mAO12 = ""
                    let mAO13 = ""
                    let mAO14 = ""
                    let mAO15 = ""
                    let mAO16 = ""
                    let mAO17 = ""
                    let mAO18 = ""
                    let mAG1 = ''
                    let mAG2 = ''
                    let mAG3 = ''
                    let mAG4 = ''
                    let mAG5 = ''
                    let mAG6 = ''
                    let mAG7 = ''
                    let mAU1 = ""
                    let mAU2 = ""
                    let mAU3 = ""
                    let mAU4 = ""
                    let mAU5 = ""
                    let mAU6 = ""
                    let mAU7 = ""
                    let mAU8 = ""
                    let mAU9 = ""
                    let mAU10 = ""
                    let mAU11 = ""
                    let mAU12 = ""
                    let mAU13 = ""
                    let mAU14 = ""
                    let mAU15 = ""
                    let mAU16 = ""
                    let mAU17 = ""
                    let mAU18 = ""
                    let mAU19 = ""
                    let mAU20 = ""
                    let mAD1 = ''
                    let mAD2 = ''
                    let mAK1 = ''
                    switch (document.getElementById('migrationAddress').value) {
                        case 'г. Москва, проспект Вернадского, 88 к. 1':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                            mAD1 = '8'; mAD2 = '8'
                            mAK1 = '1'
                            break
                        case 'г. Москва, проспект Вернадского, 88 к. 2':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                            mAD1 = '8'; mAD2 = '8'
                            mAK1 = '2'
                            break
                        case 'г. Москва, проспект Вернадского, 88 к. 3':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'П'; mAU2 = 'Р'; mAU3 = 'О'; mAU4 = 'С'; mAU5 = 'П'; mAU6 = 'Е'; mAU7 = 'К'; mAU8 = 'Т'; mAU10 = 'В'; mAU11 = 'Е'; mAU12 = 'Р'; mAU13 = 'Н'; mAU14 = 'А'; mAU15 = 'Д'; mAU16 = 'С'; mAU17 = 'К'; mAU18 = 'О'; mAU19 = 'Г'; mAU20 = 'О';
                            mAD1 = '8'; mAD2 = '8'
                            mAK1 = '3'
                            break
                        case 'г. Москва, ул. Космонавтов, д. 13':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'К'; mAU2 = 'О'; mAU3 = 'С'; mAU4 = 'М'; mAU5 = 'О'; mAU6 = 'Н'; mAU7 = 'А'; mAU8 = 'В'; mAU9 = 'Т'; mAU10 = 'О'; mAU11 = 'В';
                            mAD1 = '1'; mAD2 = '3'
                            break
                        case 'г. Москва, улица Космонавтов, д. 9':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'К'; mAU2 = 'О'; mAU3 = 'С'; mAU4 = 'М'; mAU5 = 'О'; mAU6 = 'Н'; mAU7 = 'А'; mAU8 = 'В'; mAU9 = 'Т'; mAU10 = 'О'; mAU11 = 'В';
                            mAD2 = '9'
                            break
                        case 'г. Москва, ул. Клары Цеткин, д. 25, корп. 1':
                            mAG1 = 'М'; mAG2 = 'О'; mAG3 = 'С'; mAG4 = 'К'; mAG5 = 'В'; mAG6 = 'А';
                            mAU1 = 'К'; mAU2 = 'Л'; mAU3 = 'А'; mAU4 = 'Р'; mAU5 = 'Ы'; mAU7 = 'Ц'; mAU8 = 'Е'; mAU9 = 'Т'; mAU10 = 'К'; mAU11 = 'И'; mAU12 = 'Н';
                            mAD1 = '2'; mAD2 = '5'
                            mAK1 = '1'
                            break
                        case 'Московская область, г. Люберцы, ул. Мира, д.7':
                            mAO1 = 'М'; mAO2 = 'О'; mAO3 = 'С'; mAO4 = 'К'; mAO5 = 'О'; mAO6 = 'В'; mAO7 = 'С'; mAO8 = 'К'; mAO9 = 'А'; mAO10 = 'Я'; mAO12 = 'О'; mAO13 = 'Б'; mAO14 = 'Л'; mAO15 = 'А'; mAO16 = 'С'; mAO17 = 'Т'; mAO18 = 'Ь';
                            mAG1 = 'Л'; mAG2 = 'Ю'; mAG3 = 'Б'; mAG4 = 'Е'; mAG5 = 'Р'; mAG6 = 'Ц'; mAG7 = 'Ы';
                            mAD2 = '7'
                            break
                    }


                    //numRoom {nR1-4}
                    let nR = document.getElementById('numRoom'+ indexTab).value != "-"
                        ? document.getElementById('numRoom' + indexTab).value.toUpperCase() : ''
                    let nR1 = ''
                    let nR2 = ''
                    let nR3 = ''
                    let nR4 = ''
                    if (nR) {
                        switch (nR.length) {
                            case 1:
                                nR4 = nR[0]
                                break
                            case 2:
                                nR3 = nR[0]
                                nR4 = nR[1]
                                break
                            case 3:
                                nR2 = nR[0]
                                nR3 = nR[1]
                                nR4 = nR[2]
                                break
                            case 4:
                                nR1 = nR[0]
                                nR2 = nR[1]
                                nR3 = nR[2]
                                nR4 = nR[3]
                        }
                    }






                    doc.setData({
                        lN1: lN1, lN2: lN2, lN3: lN3, lN4: lN4, lN5: lN5, lN6: lN6, lN7: lN7, lN8: lN8, lN9: lN9, lN10: lN10, lN11: lN11, lN12: lN12, lN13: lN13, lN14: lN14, lN15: lN15, lN16: lN16, lN17: lN17, lN18: lN18, lN19: lN19, lN20: lN20, lN21: lN21, lN22: lN22, lN23: lN23, lN24: lN24, lN25: lN25, lN26: lN26, lN27: lN27,
                        fN1: fN1, fN2: fN2, fN3: fN3, fN4: fN4, fN5: fN5, fN6: fN6, fN7: fN7, fN8: fN8, fN9: fN9, fN10: fN10, fN11: fN11, fN12: fN12, fN13: fN13, fN14: fN14, fN15: fN15, fN16: fN16, fN17: fN17, fN18: fN18, fN19: fN19, fN20: fN20, fN21: fN21, fN22: fN22, fN23: fN23, fN24: fN24, fN25: fN25, fN26: fN26, fN27: fN27,
                        patr1: patr1, patr2: patr2, patr3: patr3, patr4: patr4, patr5: patr5, patr6: patr6, patr7: patr7, patr8: patr8, patr9: patr9, patr10: patr10, patr11: patr11, patr12: patr12, patr13: patr13, patr14: patr14, patr15: patr15, patr16: patr16, patr17: patr17, patr18: patr18, patr19: patr19, patr20: patr20, patr21: patr21, patr22: patr22, patr23: patr23, patr24: patr24,
                        grazd1: grazd1, grazd2: grazd2, grazd3: grazd3, grazd4: grazd4, grazd5: grazd5, grazd6: grazd6, grazd7: grazd7, grazd8: grazd8, grazd9: grazd9, grazd10: grazd10, grazd11: grazd11, grazd12: grazd12, grazd13: grazd13, grazd14: grazd14, grazd15: grazd15, grazd16: grazd16, grazd17: grazd17, grazd18: grazd18, grazd19: grazd19, grazd20: grazd20, grazd21: grazd21, grazd22: grazd22, grazd23: grazd23, grazd24: grazd24, grazd25: grazd25,
                        dOB1: dOB1, dOB2: dOB2, dOB3: dOB3, dOB4: dOB4, dOB5: dOB5, dOB6: dOB6, dOB7: dOB7, dOB8: dOB8,
                        genM: genM, genW: genW,
                        pSBS1: pSBS1, pSBS2: pSBS2, pSBS3: pSBS3, pSBS4: pSBS4, pSBS5: pSBS5, pSBS6: pSBS6, pSBS7: pSBS7, pSBS8: pSBS8, pSBS9: pSBS9, pSBS10: pSBS10, pSBS11: pSBS11, pSBS12: pSBS12, pSBS13: pSBS13, pSBS14: pSBS14, pSBS15: pSBS15, pSBS16: pSBS16, pSBS17: pSBS17, pSBS18: pSBS18, pSBS19: pSBS19, pSBS20: pSBS20, pSBS21: pSBS21, pSBS22: pSBS22, pSBS23: pSBS23, pSBS24: pSBS24,
                        pSBG1: pSBG1, pSBG2: pSBG2, pSBG3: pSBG3, pSBG4: pSBG4, pSBG5: pSBG5, pSBG6: pSBG6, pSBG7: pSBG7, pSBG8: pSBG8, pSBG9: pSBG9, pSBG10: pSBG10, pSBG11: pSBG11, pSBG12: pSBG12, pSBG13: pSBG13, pSBG14: pSBG14, pSBG15: pSBG15, pSBG16: pSBG16, pSBG17: pSBG17, pSBG18: pSBG18, pSBG19: pSBG19, pSBG20: pSBG20, pSBG21: pSBG21, pSBG22: pSBG22, pSBG23: pSBG23, pSBG24: pSBG24,
                        sP1: sP1, sP2: sP2, sP3: sP3, sP4: sP4,
                        idP1: idP1, idP2: idP2, idP3: idP3, idP4: idP4, idP5: idP5, idP6: idP6, idP7: idP7, idP8: idP8, idP9: idP9, idP10: idP10,
                        dOI1: dOI1, dOI2: dOI2, dOI3: dOI3, dOI4: dOI4, dOI5: dOI5, dOI6: dOI6, dOI7: dOI7, dOI8: dOI8,
                        vU1: vU1, vU2: vU2, vU3: vU3, vU4: vU4, vU5: vU5, vU6: vU6, vU7: vU7, vU8: vU8,
                        tVV: tVV, tVJ: tVJ, tVP: tVP,
                        sV1: sV1, sV2: sV2, sV3: sV3, sV4: sV4,
                        idV1: idV1, idV2: idV2, idV3: idV3, idV4: idV4, idV5: idV5, idV6: idV6, idV7: idV7, idV8: idV8, idV9: idV9, idV10: idV10, idV11: idV11, idV12: idV12, idV13: idV13, idV14: idV14, idV15: idV15,
                        dOIV1: dOIV1, dOIV2: dOIV2, dOIV3: dOIV3, dOIV4: dOIV4, dOIV5: dOIV5, dOIV6: dOIV6, dOIV7: dOIV7, dOIV8: dOIV8,
                        vUV1: vUV1, vUV2: vUV2, vUV3: vUV3, vUV4: vUV4, vUV5: vUV5, vUV6: vUV6, vUV7: vUV7, vUV8: vUV8,
                        dAM1: dAM1, dAM2: dAM2, dAM3: dAM3, dAM4: dAM4, dAM5: dAM5, dAM6: dAM6, dAM7: dAM7, dAM8: dAM8,
                        dU1: dU1, dU2: dU2, dU3: dU3, dU4: dU4, dU5: dU5, dU6: dU6, dU7: dU7, dU8: dU8,
                        sM1: sM1, sM2: sM2, sM3: sM3, sM4: sM4,
                        iM1: iM1, iM2: iM2, iM3: iM3, iM4: iM4, iM5: iM5, iM6: iM6, iM7: iM7, iM8: iM8, iM9: iM9, iM10: iM10, iM11: iM11,
                        aHG1: aHG1, aHG2: aHG2, aHG3: aHG3, aHG4: aHG4, aHG5: aHG5, aHG6: aHG6, aHG7: aHG7, aHG8: aHG8, aHG9: aHG9, aHG10: aHG10, aHG11: aHG11, aHG12: aHG12, aHG13: aHG13, aHG14: aHG14, aHG15: aHG15, aHG16: aHG16, aHG17: aHG17, aHG18: aHG18,
                        aHU1: aHU1, aHU2: aHU2, aHU3: aHU3, aHU4: aHU4, aHU5: aHU5, aHU6: aHU6, aHU7: aHU7, aHU8: aHU8, aHU9: aHU9, aHU10: aHU10, aHU11: aHU11, aHU12: aHU12, aHU13: aHU13, aHU14: aHU14, aHU15: aHU15, aHU16: aHU16, aHU17: aHU17, aHU18: aHU18, aHU19: aHU19, aHU20: aHU20,
                        aHD1: aHD1, aHD2: aHD2, aHD3: aHD3, aHD4: aHD4, aHD5: aHD5, aHD6: aHD6, aHD7: aHD7, aHD8: aHD8, aHD9: aHD9, aHD10: aHD10, aHD11: aHD11, aHD12: aHD12, aHD13: aHD13, aHD14: aHD14, aHD15: aHD15,
                        aHK1: aHK1, aHK2: aHK2, aHK3: aHK3, aHK4: aHK4, aHK5: aHK5,
                        mAO1: mAO1, mAO2: mAO2, mAO3: mAO3, mAO4: mAO4, mAO5: mAO5, mAO6: mAO6, mAO7: mAO7, mAO8: mAO8, mAO9: mAO9, mAO10: mAO10, mAO11: mAO11, mAO12: mAO12, mAO13: mAO13, mAO14: mAO14, mAO15: mAO15, mAO16: mAO16, mAO17: mAO17, mAO18: mAO18,
                        mAG1: mAG1, mAG2: mAG2, mAG3: mAG3, mAG4: mAG4, mAG5: mAG5, mAG6: mAG6, mAG7: mAG7,
                        mAU1: mAU1, mAU2: mAU2, mAU3: mAU3, mAU4: mAU4, mAU5: mAU5, mAU6: mAU6, mAU7: mAU7, mAU8: mAU8, mAU9: mAU9, mAU10: mAU10, mAU11: mAU11, mAU12: mAU12, mAU13: mAU13, mAU14: mAU14, mAU15: mAU15, mAU16: mAU16, mAU17: mAU17, mAU18: mAU18, mAU19: mAU19, mAU20: mAU20,
                        mAD1: mAD1, mAD2: mAD2, mAK1: mAK1,


                        purp1: purp1, purp2: purp2, purp3: purp3, purp4: purp4, purp5: purp5, purp6: purp6, purp7: purp7, purp8: purp8, purp9: purp9, purp10: purp10, purp11: purp11, purp12: purp12, purp13: purp13,
                        purposeG: purposeG, purposeU: purposeU, purposeR: purposeR,

                        nR1: nR1, nR2: nR2, nR3: nR3, nR4: nR4,


                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(

                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + "/"+
                        "УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         +".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" УВЕДОМЛЕНИЕ (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         + ".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                // saveAs(content,nameFile);
            });
    };

//ходатайство РЕГИСТРАЦИЯ
    window.generateRegSolicitaionTot = function generate() {
        path = ("")
        switch (document.getElementById('ovmByRegion').value) {
            case "Алексеевский":
                path = ('../Templates/регистрация/ходатайство АЛЕКСЕЕВСКИЙ.docx')
                break
            case "Войковский":
                path = ('../Templates/регистрация/ходатайство ВОЙКОВСКИЙ.docx')
                break
            case "МУ МВД РФ Люберецкое":
                path = ('../Templates/регистрация/ходатайство МУ МВД РФ ЛЮБЕРЕЦКОЕ.docx')
                break
            case "Тропарево-Никулино":
                path = ('../Templates/регистрация/ходатайство ТРОПАРЕВО-НИКУЛИНО.docx')
                break
        }

        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))


                    //dateUntil
                    let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                    // purpose
                    let purpose = ''
                    let purposeS = ''
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purpose = "УЧЕБА"
                            purposeS = "Студент"
                            break
                        case "Краткосрочная учеба":
                            purpose = "КРАТКОСРОЧНАЯ УЧЕБА"
                            purposeS = "Студент"
                            break
                        case "(НТС)":
                            purpose = "НАУЧНО-ТЕХНИЧЕСКИЕ СВЯЗИ (НТС)"
                            purposeS = "НТС"
                            break
                        case "Трудовая деятельность":
                            purpose = "ТРУДОВАЯ ДЕЯТЕЛЬНОСТЬ"
                            purposeS = "Преподаватель"
                            break
                    }

                    // levelEducation
                    let levelEducation = ''
                    switch (document.getElementById('levelEducation' + indexTab).value) {
                        case "Подготовительный факультет (изучаю русский язык)/ The preparatory faculty":
                            levelEducation = 'подготовительный факультет'
                            break
                        case "бакалавриат/bachelor degree":
                            levelEducation = 'бакалавриат'
                            break
                        case "магистратура/master degree":
                            levelEducation = 'магистратура'
                            break
                        case "аспирантура/post-graduate studies":
                            levelEducation = 'аспирантура'
                            break
                    }

                    // course
                    let course = ''
                    switch (document.getElementById('course' + indexTab).value) {
                        case '1':
                            course = ', 1 курс,'
                            break
                        case '2':
                            course = ', 2 курс,'
                            break
                        case '3':
                            course = ', 3 курс,'
                            break
                        case '4':
                            course = ', 4 курс,'
                            break
                        case '5':
                            course = ', 5 курс,'
                            break
                    }

                    // addressResidence
                    let addressResidence = ''
                    switch (document.getElementById('migrationAddress').value) {
                        case "Квартира":
                            addressResidence = document.getElementById('addressResidence' + indexTab).value
                            break
                        default:
                            addressResidence = document.getElementById('migrationAddress').value
                            break
                    }

                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                    // gender
                    let gender = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            gender = 'м.'
                            break
                        case "Женский / Female":
                            gender = 'ж.'
                            break
                    }

                    // registration On
                    let registrationOn = ''
                    switch (document.getElementById('registrationOn').value) {
                        case "Круглов":
                            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                            break
                        case "Морозова":
                            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                            break
                        case "Орлова":
                            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                            break
                    }

                    // visa
                    let typeVisa = ''
                    switch (document.getElementById('typeVisa' + indexTab).value) {
                        case "ВИЗА":
                            typeVisa = 'Виза'
                            break
                        case "(ВНЖ) ВИД НА ЖИТЕЛЬСТВО РФ":
                            typeVisa = 'ВНЖ'
                            break
                        case "(РВП) РАЗРЕШЕНИЕ НА ВРЕМЕННОЕ ПРОЖИВАНИЕ РФ":
                            typeVisa = 'РВП'
                            break

                    }
                    let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value : ''
                    let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                        ? document.getElementById('idVisa' + indexTab).value : ''

                    let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                    let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                    // order
                    let numOrder = document.getElementById('numOrder' + indexTab).value
                        ? document.getElementById('numOrder' + indexTab).value : ''
                    let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''
                    let orderUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString() : ''

                    // faculty
                    let faculty = ''
                    switch (document.getElementById('faculty' + indexTab).value) {
                        case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                            faculty = 'ИИИ:Музфак'
                            break
                        case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                            faculty = 'ИИИ: Худграф'
                            break
                        case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                            faculty = 'ИСГО'
                            break
                        case "Институт филологии / The Institute of Philology":
                            faculty = 'ИФ'
                            break
                        case "Институт иностранных языков / The Institute of Foreign Languages":
                            faculty = 'ИИЯ'
                            break
                        case "Институт международного образования / The Institute of International Education":
                            faculty = 'ИМО'
                            break
                        case "Институт детства / The Institute of Childhood":
                            faculty = 'ИД'
                            break
                        case "Институт биологии и химии / The Institute of Biology and Chemistry":
                            faculty = 'ИБХ'
                            break
                        case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                            faculty = 'ИФТИС'
                            break
                        case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                            faculty = 'ИФКСиЗ'
                            break
                        case "Географический факультет / The Institute of Geography":
                            faculty = 'Геофак'
                            break
                        case "Институт истории и политики / The Institute of History and Politics":
                            faculty = 'ИИП'
                            break
                        case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                            faculty = 'ИМИ'
                            break
                        case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                            faculty = 'Дош.фак.'
                            break
                        case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                            faculty = 'ИПП'
                            break
                        case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                            faculty = 'ИЖКиМ'
                            break
                        case "Институт развития цифрового образования / The Institute of Digital Education Development":
                            faculty = 'ИРЦО'
                            break
                    }


                    // contract
                    let typeFundingDog1 = ""
                    let typeFundingDog2 = ""
                    let typeFundingNap1 = ""
                    let typeFundingNap2 = ""
                    switch (document.getElementById('typeFunding' + indexTab).value) {
                        case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                            typeFundingDog2 = "Договор"
                            typeFundingNap2 = "направление"
                            break
                        case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                            typeFundingDog1 = "Договор"
                            typeFundingNap1 = "направление"
                            break
                    }
                    let numContract = document.getElementById('numContract' + indexTab).value
                        ? document.getElementById('numContract' + indexTab).value : ''
                    let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                        document.getElementById('contractFrom' + indexTab).value : ''


                    let numRoom = document.getElementById('numRoom' + indexTab).value != "" ? ', комната № ' + document.getElementById('numRoom' + indexTab).value : ''

                    let numRental = document.getElementById('numRental' + indexTab).value != "-" ? document.getElementById('numRental' + indexTab).value : ''


                    doc.setData({
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                        nStud: document.getElementById('nStud' + indexTab).value,
                        grazd: document.getElementById('grazd' + indexTab).value,
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                        dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                        gender: gender,

                        registrationOn: registrationOn,
                        dateUntil: dateUntil,
                        purpose: purpose,
                        purposeS: purposeS,
                        levelEducation: levelEducation,
                        course: course,

                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,

                        typeVisa: typeVisa,
                        seriesVisa: seriesVisa,
                        idVisa: idVisa,
                        dateOfIssueVisa: dateOfIssueVisa,
                        validUntilVisa: validUntilVisa,

                        seriesMigration: document.getElementById('seriesMigration' + indexTab).value,
                        idMigration: document.getElementById('idMigration' + indexTab).value,
                        dateArrivalMigration: new Date(document.getElementById('dateArrivalMigration' + indexTab).value).toLocaleDateString(),



                        migrationAddress: document.getElementById('migrationAddress').value,
                        numRoom: '',
                        faculty: faculty,
                        numOrder: numOrder,
                        orderFrom: orderFrom,
                        orderUntil: orderUntil,

                        typeFundingDog1: typeFundingDog1,
                        typeFundingDog2: typeFundingDog2,
                        typeFundingNap1: typeFundingNap1,
                        typeFundingNap2: typeFundingNap2,

                        numContract: numContract,
                        contractFrom: contractFrom,

                        numRental: '',

                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + "/" +

                        "ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         +".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" ХОДАТАЙСТВО (РЕГИСТРАЦИЯ) - " +
                //         document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                //         + ".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                //saveAs(content,nameFile);
            });
    };

//визовая анкета
    window.generateVisaApplicationTot = function generate() {
        path = ('../Templates/виза/визовая анкета.docx')
        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))

                    // purpose
                    let purposeG = ""
                    let purposeR = ""
                    let purposeU = ""
                    let purposeS = "-"
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purposeU = "X"
                            purposeS = "Студент"
                            break
                        case "Краткосрочная учеба":
                            purposeU = "X"
                            purposeS = "Студент"
                            break
                        case "(НТС)":
                            purposeG = "X"
                            break
                        case "Трудовая деятельность":
                            purposeR = "X"
                            purposeS = "Преподаватель"
                            break
                    }

                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                    // gender
                    let genM = ''
                    let genW = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            genM = 'X'
                            break
                        case "Женский / Female":
                            genW = 'X'
                            break
                    }

                    // visa
                    let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value : ''
                    let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                        ? document.getElementById('idVisa' + indexTab).value : ''
                    let identifierVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('identifierVisa' + indexTab).value)
                        ? document.getElementById('identifierVisa' + indexTab).value : ''
                    let numInvVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('numInvVisa' + indexTab).value)
                        ? document.getElementById('numInvVisa' + indexTab).value : ''

                    let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                    let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''


                    // OVM
                    let infHost1 = ''
                    let infHost2 = ''
                    let addressResidence = ''
                    let numRoom = ''
                    switch (document.getElementById('migrationAddress').value) {
                        case "Квартира":
                            if (document.getElementById('infHost' + indexTab).value.length>60) {
                                let totalLen = 0
                                let lenInfHost = document.getElementById('infHost' + indexTab).value.split(',')
                                for (let i = 0; i< lenInfHost.length; i++) {
                                    totalLen = totalLen + lenInfHost[i].length
                                    if (totalLen<60) {
                                        if (i==lenInfHost.length-1) {infHost1 = infHost1  + lenInfHost[i]}
                                        else {infHost1 = infHost1  + lenInfHost[i] + ', '}
                                    }
                                    else {
                                        if (i==lenInfHost.length-1) {infHost2 = infHost2  + lenInfHost[i]}
                                        else {infHost2 = infHost2  + lenInfHost[i] + ', '}
                                    }
                                }
                            }
                            else {infHost1 = document.getElementById('infHost' + indexTab).value}
                            addressResidence = document.getElementById('addressResidence' + indexTab).value
                            break
                        default:
                            infHost1 = 'МПГУ, 7704077771, М. Пироговская д. 1, стр. 1.'
                            infHost2 = '8-499-245-03-10, mail@mpgu.su'
                            addressResidence = document.getElementById('migrationAddress').value
                            numRoom = ', комната № ' + document.getElementById('numRoom' + indexTab).value
                            break

                    }

                    doc.setData({


                        purposeG: purposeG,
                        purposeR: purposeR,
                        purposeU: purposeU,
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                        lastNameEn: document.getElementById('lastNameEn' + indexTab).value.toUpperCase(),
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value.toUpperCase(),
                        firstNameEn: document.getElementById('firstNameEn' + indexTab).value.toUpperCase(),
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value.toUpperCase(),
                        dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                        placeStateBirth: document.getElementById('placeStateBirth' + indexTab).value.toUpperCase(),
                        genM: genM,
                        genW: genW,
                        grazd: document.getElementById('grazd' + indexTab).value,
                        // documentPerson: document.getElementById('documentPerson' + indexTab).text,
                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,

                        infHost1: infHost1,
                        infHost2: infHost2,
                        addressResidence: addressResidence,
                        numRoom: '',

                        homeAddress: document.getElementById('homeAddress' + indexTab).value,
                        purposeS: purposeS,
                        phone: document.getElementById('phone' + indexTab).value,
                        mail: document.getElementById('mail' + indexTab).value,
                        seriesVisa: seriesVisa,
                        idVisa: idVisa,
                        identifierVisa: identifierVisa,
                        dateOfIssueVisa: dateOfIssueVisa,
                        validUntilVisa: validUntilVisa,
                        numInvVisa: numInvVisa,
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),

                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + "/" +

                        "ВИЗОВАЯ АНКЕТА - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" ВИЗОВАЯ АНКЕТА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text +".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" ВИЗОВАЯ АНКЕТА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text +".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                //saveAs(content,nameFile);
            });
    };

//ходатайство ВИЗА ТРОПАРЕВО-НИКУЛИНО
    window.generateVisaSolicitaionTroparevoTot = function generate() {
        path = ('../Templates/виза/ходатайство ТРОПАРЕВО-НИКУЛИНО.docx')

        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))


                    //dateUntil
                    let dateUntil= /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

                    // purpose
                    let purpose = ''
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purpose = "обучением в МПГУ"
                            break
                        case "Краткосрочная учеба":
                            purpose = "обучением в МПГУ"
                            break
                        case "(НТС)":
                            purpose = "посещением МПГУ в качестве приглашенного гостя (НТС)"
                            break
                        case "Трудовая деятельность":
                            purpose = "преподавательской деятельностью в МПГУ"
                            break
                    }

                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''

                    // gender
                    let genM = ''
                    let genW = ''
                    switch (document.getElementById('gender' + indexTab).value) {
                        case "Мужской / Male":
                            genM = 'X'
                            genW = ' '
                            break
                        case "Женский / Female":
                            genW = 'X'
                            genM = ' '
                            break
                    }

                    // registration On
                    let registrationOn1 = ''
                    let registrationOn2 = ''
                    switch (document.getElementById('registrationOn').value) {
                        case "Круглов":
                            registrationOn1 = 'Начальник УМС                                                    Круглов В.В.'
                            registrationOn2 = 'Начальник УМС                                                                          В. В. Круглов'
                            break
                        case "Морозова":
                            registrationOn1 = 'Заместитель начальника УМС                Морозова О.А.'
                            registrationOn2 = 'Заместитель начальника УМС                                                Морозова О.А.'
                            break
                        case "Орлова":
                            registrationOn1 = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                            registrationOn2 = 'Начальник паспортно-визового отдела УМС                              Орлова С.В.'
                            break
                    }

                    // visa
                    let seriesVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('seriesVisa' + indexTab).value)
                        ? document.getElementById('seriesVisa' + indexTab).value : ''
                    let idVisa = /^[a-zA-Z0-9.]+$/.test(document.getElementById('idVisa' + indexTab).value)
                        ? document.getElementById('idVisa' + indexTab).value : ''

                    let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''
                    let validUntilVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntilVisa' + indexTab).value).toLocaleDateString() : ''

                    // order
                    let numOrder = document.getElementById('numOrder' + indexTab).value
                        ? document.getElementById('numOrder' + indexTab).value : ''
                    let orderFrom = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderFrom' + indexTab).value).toLocaleDateString() : ''
                    let orderUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('orderUntil' + indexTab).value).toLocaleDateString() : ''

                    // contract
                    let typeFunding = ''
                    switch (document.getElementById('typeFunding' + indexTab).value) {
                        case "бюджет (Государство оплачивает мое обучение)/ state funded (The state pays for my education)":
                            typeFunding = 'НАПРАВЛЕНИЕ №'
                            break
                        case 'договор ( я плачу за обучение)/paid tuition (I pay for my education)':
                            typeFunding = 'ДОГОВОР №'
                            break
                    }
                    let numContract = document.getElementById('numContract' + indexTab).value
                        ? document.getElementById('numContract' + indexTab).value : ''
                    let contractFrom = document.getElementById('contractFrom' + indexTab).value != '-' ?
                        document.getElementById('contractFrom' + indexTab).value : ''



                    doc.setData({
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                        nStud: document.getElementById('nStud' + indexTab).value,
                        grazd: document.getElementById('grazd' + indexTab).value,
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value,
                        lastNameEn: document.getElementById('lastNameEn' + indexTab).value,
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value,
                        firstNameEn: document.getElementById('firstNameEn' + indexTab).value,
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value,
                        dateOfBirth: new Date(document.getElementById('dateOfBirth' + indexTab).value).toLocaleDateString(),
                        registrationOn1: registrationOn1,
                        registrationOn2: registrationOn2,
                        dateUntil: dateUntil,
                        genM: genM,
                        genW: genW,
                        purpose: purpose,

                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,

                        seriesVisa: seriesVisa,
                        idVisa: idVisa,
                        dateOfIssueVisa: dateOfIssueVisa,
                        validUntilVisa: validUntilVisa,


                        numOrder: numOrder,
                        orderFrom: orderFrom,
                        orderUntil: orderUntil,

                        typeFunding: typeFunding,
                        numContract: numContract,
                        contractFrom: contractFrom,

                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + "ОВМ ТРОПАРЕВО-НИКУЛИНО" + "/" +
                        "ХОДАТАЙСТВО (ВИЗА) - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + "ОВМ ТРОПАРЕВО-НИКУЛИНО" + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ТРОПАРЕВО-НИКУЛИНО.zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" ХОДАТАЙСТВО (ВИЗА) - ОВМ ТРОПАРЕВО-НИКУЛИНО.zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                //saveAs(content,nameFile);
            });
    };

    //справка
    window.generateVisaReferenceTot = function generate() {
        path = ('../Templates/виза/справка.docx')
        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }

                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                function errorHandler(error) {
                    console.log(JSON.stringify({error: error}, replaceErrors));

                    if (error.properties && error.properties.errors instanceof Array) {
                        const errorMessages = error.properties.errors.map(function (error) {
                            return error.properties.explanation;
                        }).join("\n");
                        console.log('errorMessages', errorMessages);
                        // errorMessages is a humanly readable message looking like this :
                        // 'The tag beginning with "foobar" is unopened'
                    }
                    throw error;
                }

                for (let i =0; i<countTab();i++) {
                    var zip = new PizZip(content);
                    var doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });

                    let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
                    let elem = tabs[i]
                    let indexTab = parseInt(elem.id.match(/\d+/))

                    // purpose
                    let purpose = ""
                    switch (document.getElementById('purpose' + indexTab).value) {
                        case "Учеба":
                            purpose = "студентом"
                            break
                        case "Краткосрочная учеба":
                            purpose = "студентом"
                            break
                        case "(НТС)":
                            purpose = "приглашенным гостем (НТС)"
                            break
                        case "Трудовая деятельность":
                            purpose = "преподавателем"
                            break
                    }

                    // faculty
                    let faculty = ''
                    switch (document.getElementById('faculty' + indexTab).value) {
                        case "Институт изящных искусств: Факультет музыкального искусства / The Musical Arts Institute":
                            faculty = 'Института изящных искусств: Факультета музыкального искусства'
                            break
                        case "Институт изящных искусств: Художественно-графический факультет/ The Institute of Fine Arts":
                            faculty = 'Института изящных искусств: Художественно-графического факультета'
                            break
                        case "Институт социально-гуманитарного образования / The Institute of Social Studies and Humanities":
                            faculty = 'Института социально-гуманитарного образования'
                            break
                        case "Институт филологии / The Institute of Philology":
                            faculty = 'Института филологии'
                            break
                        case "Институт иностранных языков / The Institute of Foreign Languages":
                            faculty = 'Института иностранных языков'
                            break
                        case "Институт международного образования / The Institute of International Education":
                            faculty = 'Института международного образования'
                            break
                        case "Институт детства / The Institute of Childhood":
                            faculty = 'Института детства'
                            break
                        case "Институт биологии и химии / The Institute of Biology and Chemistry":
                            faculty = 'Института биологии и химии'
                            break
                        case "Институт физики, технологии и информационных систем / The Institute of Physics, Technology, and Informational Systems":
                            faculty = 'Института физики, технологии и информационных систем'
                            break
                        case "Институт физической культуры, спорта и здоровья /The Institute of Physical Education, Sports and Health":
                            faculty = 'Института физической культуры, спорта и здоровья'
                            break
                        case "Географический факультет / The Institute of Geography":
                            faculty = 'Географического факультета'
                            break
                        case "Институт истории и политики / The Institute of History and Politics":
                            faculty = 'Института истории и политики'
                            break
                        case "Институт математики и информатики / The Institute of Mathematics and Informatics":
                            faculty = 'Института математики и информатики'
                            break
                        case "Факультет дошкольной педагогики и психологии / The Institute of Pre-School Pedagogy and Psychology":
                            faculty = 'Факультета дошкольной педагогики и психологии'
                            break
                        case "Институт педагогики и психологии / The Institute of Pedagogy and Psychology":
                            faculty = 'Института педагогики и психологии'
                            break
                        case "Институт журналистики, коммуникаций и медиаобразования / The Institute of Journalism, Communications and Media Education":
                            faculty = 'Института журналистики, коммуникаций и медиаобразования'
                            break
                        case "Институт развития цифрового образования / The Institute of Digital Education Development":
                            faculty = 'Института развития цифрового образования'
                            break

                    }


                    // Passport
                    let series = /^[a-zA-Z0-9.]+$/.test(document.getElementById('series' + indexTab).value)
                        ? document.getElementById('series' + indexTab).value : ''
                    let validUntil = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString())
                        ? new Date(document.getElementById('validUntil' + indexTab).value).toLocaleDateString() : ''



                    // OVM
                    let ovmByRegion = ''
                    switch (document.getElementById('ovmByRegion').value) {
                        case "Тропарево-Никулино":
                            ovmByRegion = 'ОМВД России по району Тропарево-Никулино г.Москвы'
                            break
                        case "Хамовники":
                            ovmByRegion = 'ОМВД России по району Хамовники г.Москвы'
                            break
                    }

                    // registration On
                    let registrationOn = ''
                    switch (document.getElementById('registrationOn').value) {
                        case "Круглов":
                            registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                            break
                        case "Морозова":
                            registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                            break
                        case "Орлова":
                            registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                            break
                    }

                    let dateUnt =  document.getElementById('dateUntil' + indexTab).value ? new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''


                    doc.setData({

                        grazd: document.getElementById('grazd' + indexTab).value.toUpperCase(),
                        lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                        firstNameRu: document.getElementById('firstNameRu' + indexTab).value.toUpperCase(),
                        patronymicRu: document.getElementById('patronymicRu' + indexTab).value.toUpperCase(),
                        purpose: purpose,
                        faculty: faculty.toUpperCase(),
                        series: series,
                        idPassport: document.getElementById('idPassport' + indexTab).value,
                        dateOfIssue: new Date(document.getElementById('dateOfIssue' + indexTab).value).toLocaleDateString(),
                        validUntil: validUntil,
                        dateUntil: dateUnt,
                        ovmByRegion: ovmByRegion,
                        registrationOn: registrationOn,
                        nStud: document.getElementById('nStud' + indexTab).value,
                        dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    });



                    try {
                        doc.render();
                    }
                    catch (error) {
                        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                        errorHandler(error);
                    }
                    var out = doc.getZip().generate();
                    zipTotal.file(
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + "/" +

                        "СПРАВКА - (" + document.getElementById('grazd'+indexTab).value.toUpperCase() + ") " +
                        document.getElementById('lastNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('firstNameRu'+indexTab).value.toUpperCase() + ' ' +
                        document.getElementById('patronymicRu'+indexTab).value.toUpperCase() + " - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString() +
                        " - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".docx"
                        , out, {base64: true}
                    );
                } // end for


                // let nameFile = ''
                // if (countTab()==1) {
                //     nameFile = document.getElementById('nStud1').value
                //         +" СПРАВКА - "  + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".zip"
                // }
                // else {
                //     nameFile = document.getElementById('nStud1').value + '-'+
                //         document.getElementById('nStud'+(lastTab()-1)).value
                //         +" СПРАВКА - " + document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text + ".zip"
                // }
                var content = zipTotal.generate({ type: "blob" });
                // saveAs(content,nameFile);
            });
    };


    //опись
    window.generateInventoryRegTot = function generate() {
        path = ('../Templates/регистрация И виза/опись.docx')

        let students = []

        // registration On
        let registrationOn = ''
        switch (document.getElementById('registrationOn').value) {
            case "Круглов":
                registrationOn = 'Начальник УМС                                                    Круглов В.В.'
                break
            case "Морозова":
                registrationOn = 'Заместитель начальника УМС                Морозова О.А.'
                break
            case "Орлова":
                registrationOn = 'Начальник паспортно-визового отдела УМС Орлова С.В.'
                break
        }


        // nStud
        let nStud1 = document.getElementById('nStud1').value
        let tir = ''
        let nStud2 = ''
        if (countTab()>1) {
            nStud2 = document.getElementById('nStud' + (lastTab()-1)).value
            tir = '-'
        }

        for (let i =0; i<countTab();i++) {
            let tabs = document.getElementsByClassName('nav-tabs')[0].getElementsByTagName('li')
            let elem = tabs[i]
            let indexTab = parseInt(elem.id.match(/\d+/))


            let dateOfIssueVisa = /^[a-zA-Z0-9.]+$/.test(new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString())
                ? new Date(document.getElementById('dateOfIssueVisa' + indexTab).value).toLocaleDateString() : ''

            let dateUntil = document.getElementById('dateUntil' + indexTab).value != '-' ?
                new Date(document.getElementById('dateUntil' + indexTab).value).toLocaleDateString() : ''

            //students
            students.push({
                nStud: document.getElementById('nStud' + indexTab).value,
                lastNameRu: document.getElementById('lastNameRu' + indexTab).value.toUpperCase(),
                firstNameRu: document.getElementById('firstNameRu'+indexTab).value.toUpperCase(),
                patronymicRu: document.getElementById('patronymicRu'+indexTab).value.toUpperCase(),
                dateOfIssueVisa: dateOfIssueVisa,
                dateUntil: dateUntil,
                grazd: document.getElementById('grazd' + indexTab).value,
                phone: document.getElementById('phone' + indexTab).value,
                mail: document.getElementById('mail' + indexTab).value,
            })
        }


        loadFile(
            path,
            function (error, content) {
                if (error) {
                    throw error;
                }
                var zip = new PizZip(content);
                var doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                });

                doc.render({
                    dateInOvm: new Date(document.getElementById('dateInOvm').value).toLocaleDateString(),
                    nStud1: nStud1,
                    nStud2: nStud2,
                    tir: tir,
                    'students': students,
                    registrationOn: registrationOn,
                })


                var out = doc.getZip().generate();



                let nameFile = ''
                if (countTab()==1) {
                    nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ + ВИЗА) - студент " +
                        document.getElementById('nStud1').value
                        +" - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                        +" - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                }
                else {
                    nameFile = "ОПИСЬ (РЕГИСТРАЦИЯ + ВИЗА) - студенты " +
                        document.getElementById('nStud1').value
                        +"-"+
                        document.getElementById('nStud'+(lastTab()-1)).value
                        +" - " +
                        new Date(document.getElementById('dateInOvm').value).toLocaleDateString()
                        +" - " +
                        document.getElementById('ovmByRegion').options[document.getElementById('ovmByRegion').selectedIndex].text
                        + ".docx"
                }




                zipTotal.file(nameFile, out, {base64: true})

                // Output the document using Data-URI
                //saveAs(out, nameFile);
                var content = zipTotal.generate({type: 'blob'})

                let nameZip = ''
                if (countTab()==1) {
                    nameZip = document.getElementById('nStud1').value
                        +" Студент.zip"
                }
                else {
                    nameZip = document.getElementById('nStud1').value + '-'+
                        document.getElementById('nStud'+(lastTab()-1)).value
                        +" Студенты.zip"
                }

                saveAs(content, nameZip)
            }
        );
    };



    setTimeout(generateRegNotifTot, 30)
    setTimeout(generateRegSolicitaionTot, 30)
    setTimeout(generateVisaApplicationTot, 30)
    setTimeout(generateVisaSolicitaionTroparevoTot, 30)
    setTimeout(generateVisaReferenceTot, 30)
    setTimeout(generateInventoryRegTot, 300)
}








