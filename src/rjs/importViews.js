const { dialog } = require('@electron/remote');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

class ImportRosterView extends View {
    get name() { return 'Roster'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Import Roster'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'From here, you can import the Missionary Portal roster spreadsheet (.xlsx). In no particular order, you should include the fields for =Missionary=, =ID=, =Type=, =Position Abbr=, =Phone=, =Assignment=, =Status=, =Arrival Date=, =Release Date=, =Area=, =District=, and =Zone=.<br>Make sure you create a custom view of the roster in Missionary Portal to ensure you export each column necessary!<br>In order to do that:<br><ol><li></li>'
        });

        this.addElement(comment);

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Select File',
            eventListener: ['click', () => {
                const path = dialog.showOpenDialogSync({
                    properties: ['openFile'],
                    filters: [ { name: 'Excel Spreadsheet (.xlsx)', extensions: ['xlsx'] } ]
                });

                if (path) {
                    const file = fs.readFileSync(path[0]);
                    const workbook = XLSX.read(file);

                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const data = XLSX.utils.sheet_to_json(firstSheet);

                    this.database.importTable(data);
                }
            }]
        });

        this.addElement(button);
    }
}

class ImportAddressesView extends View {
    get name() { return 'Addresses'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Import Addresses'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'From here, you can import a spreadsheet containing apartment addresses (.xlsx). In no particular order, you should include fields for =Name= (a label for the apartment), =Postal Code=, =Japanese Address=, and =English Address=.'
        });

        this.addElement(comment);

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Select File',
            eventListener: ['click', () => {
                const path = dialog.showOpenDialogSync({
                    properties: ['openFile'],
                    filters: [ { name: 'Excel Spreadsheet (.xlsx)', extensions: ['xlsx'] } ]
                });

                if (path) {
                    const file = fs.readFileSync(path[0]);
                    const workbook = XLSX.read(file);

                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const data = XLSX.utils.sheet_to_json(firstSheet);

                    this.database.importAddresses(data);
                }
            }]
        });

        this.addElement(button);
    }
}

class ImportAddressKeyView extends View {
    get name() { return 'Address Key'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Import Address Key'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'From here, you can import a spreadsheet that associates apartment names with area names (.xlsx). In no particular order, you should include fields for =Areas= and =Apartment=. Each row should have one area pointing to one apartment label.'
        });

        this.addElement(comment);

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Select File',
            eventListener: ['click', () => {
                const path = dialog.showOpenDialogSync({
                    properties: ['openFile'],
                    filters: [ { name: 'Excel Spreadsheet (.xlsx)', extensions: ['xlsx'] } ]
                });

                if (path) {
                    const file = fs.readFileSync(path[0]);
                    const workbook = XLSX.read(file);

                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const data = XLSX.utils.sheet_to_json(firstSheet);

                    this.database.importAddressKey(data);
                }
            }]
        });

        this.addElement(button);
    }
}

class ImportProfilesView extends View {
    get name() { return 'Profiles'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Import Profiles'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'Use this to import many missionary profile pictures at once. The images should be named =ID.jpg= or =ID.txt=, where "ID" is the name of the missionary. Since we will try compressing each image, this might take a few minutes.'
        });

        this.addElement(comment);

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Select Folder',
            eventListener: ['click', () => {
                const folderPath = dialog.showOpenDialogSync({
                    properties: ['openDirectory']
                });

                if (folderPath) {
                    let data = {};

                    for (const fileName of fs.readdirSync(folderPath[0])) {
                        const extension = path.extname(fileName).slice(1);
                        const ID = path.basename(fileName, '.' + extension);

                        if (extension !== 'txt' && extension !== 'jpg' || !/^\d{6}$/.test(ID)) {
                            continue;
                        }

                        const filePath = path.join(folderPath[0], fileName);
                        
                        data[ID] = {
                            filePath: filePath,
                            extension: extension,
                            id: ID
                        };
                    }

                    this.database.importProfiles(data);
                }
            }]
        });

        this.addElement(button);
    }
}

class ImportCoversView extends View {
    get name() { return 'Covers'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Import Covers'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'Use this to import several zone cover pictures at once. The images should be named =zone.png= or =zone.txt=, where "zone" is the name of the zone.'
        });

        this.addElement(comment);

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Select Folder',
            eventListener: ['click', () => {
                const folderPath = dialog.showOpenDialogSync({
                    properties: ['openDirectory']
                });

                if (folderPath) {
                    let data = {};

                    for (const fileName of fs.readdirSync(folderPath[0])) {
                        const extension = path.extname(fileName).slice(1);
                        const ID = path.basename(fileName, '.' + extension);

                        if (extension !== 'txt' && extension !== 'jpg' && extension !== 'png') {
                            continue;
                        }

                        const filePath = path.join(folderPath[0], fileName);
                        
                        data[ID] = {
                            filePath: filePath,
                            extension: extension,
                            id: ID
                        };
                    }

                    this.database.importCovers(data);
                }
            }]
        });

        this.addElement(button);
    }
}