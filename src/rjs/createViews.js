const { spawn } = require('child_process');

class CreateTransferDocsView extends View {
    get name() { return 'Transfer Docs'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Generate Transfer Docs'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'Use this to generate a PDF copy of transfer docs including cover pictures, phone numbers, addresses, area and missionary information, and the transfer calendar.'
        });

        this.addElement(comment);

        const zoneAreas = this.database.getZoneAreas();
        const covers = this.database.getCovers();

        const coverGallery = new Element('DIV', null, {
            elementClass: 'view-double-gallery'
        });

        let names = Object.keys(zoneAreas);
        names.splice(0, 0, 'Cover');
        names.push('Transfer Board');
        names.push('Calendar');

        for (const name of names) {
            const galleryEntry = new Element('DIV', coverGallery, {
                elementClass: 'view-gallery-entry'
            });

            new Element('H2', galleryEntry, {
                elementClass: 'view-gallery-header',
                text: name
            });

            const image = new Element('IMG', galleryEntry, {
                elementClass: 'view-gallery-picture',
                attributes: {
                    'SRC': covers[name]
                }
            });

            new Element('BUTTON', galleryEntry, {
                elementClass: 'view-gallery-button',
                text: 'Upload',
                eventListener: ['click', () => {
                    const filePath = dialog.showOpenDialogSync({
                        properties: ['openFile'],
                        filters: [
                            { name: 'Image', extensions: ['jpg', 'jpeg', 'png'] }
                        ]
                    });

                    if (filePath) {
                        const raw = this.database.importCover(
                            filePath[0],
                            name,
                            name === 'Transfer Board' || name === 'Calendar'
                        );

                        image.element.src = raw;
                    }
                }]
            });
        }

        this.addElement(coverGallery);

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Generate',
            eventListener: ['click', () => {
                ipcRenderer.sendSync('create-window', 'generate', {
                    show: false,
                    webPreferences: {
                        nodeIntegration: true,
                        enableRemoteModule: true,
                        contextIsolation: false
                    },
                    titleBarStyle: 'hidden',
                    titleBarOverlay: {
                        color: '#292929',
                        symbolColor: '#fff',
                    },
                    minWidth: 600,
                    minHeight: 450,
                    backgroundColor: '#fff',
                });
            }]
        });

        this.addElement(button);
    }
}

function encodeQuotedPrintable(input) {
    let output = '';

    for (var i = 0; i < input.length; i++) {
        var charCode = input.charCodeAt(i);
        if (charCode === 61) { // '='
            output += '=3D';
        } else if (charCode < 32 || charCode > 126) {
            output += '=' + charCode.toString(16).toUpperCase().padStart(2, '0');
        } else {
            output += String.fromCharCode(charCode);
        }
    }
    return output;
}

function generateCard(name, reading, number, missionaries, photo, note, zone, group) {
    const template = `
BEGIN:VCARD
VERSION:2.1
FN:${name}
SOUND;X-IRMC-N;CHARSET=UTF-8:${reading};;;;
TEL;CELL:${number}
TEL;CELL:
TEL;CELL:
TEL;CELL:
EMAIL;HOME:
ADR;HOME:;;;;;;
ORG:
TITLE:${missionaries}
PHOTO;ENCODING=BASE64;PNG:${photo}
NOTE;ENCODING=QUOTED-PRINTABLE:${note}
X-GN:${group}
X-CLASS:PUBLIC
X-REDUCTION:
X-NO:
X-DCM-HMN-MODE:
END:VCARD
`;

    return template;
}

class CreateContactsView extends View {
    get name() { return 'Contacts'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Generate Contacts Card'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'Use this to generate a =.vcf= file containing contact cards for each area.'
        });

        this.addElement(comment);

        const zoneAreas = this.database.getZoneAreas();
        const contactFaces = this.database.getContactFaces();

        const faceGallery = new Element('DIV', null, {
            elementClass: 'view-double-gallery'
        });

        let groups = [];

        for (const number of Object.values(this.database.numbers)) {
            if (groups.indexOf(number.group) === -1) {
                groups.push(number.group);
            }
        }

        for (const name of groups) {
            const galleryEntry = new Element('DIV', faceGallery, {
                elementClass: 'view-gallery-entry'
            });

            new Element('H2', galleryEntry, {
                elementClass: 'view-gallery-header',
                text: name
            });

            const image = new Element('IMG', galleryEntry, {
                elementClass: 'view-gallery-picture',
                attributes: {
                    'SRC': contactFaces[name]
                }
            });

            new Element('BUTTON', galleryEntry, {
                elementClass: 'view-gallery-button',
                text: 'Upload',
                eventListener: ['click', () => {
                    const filePath = dialog.showOpenDialogSync({
                        properties: ['openFile'],
                        filters: [
                            { name: 'Image', extensions: ['jpg', 'jpeg', 'png'] }
                        ]
                    });

                    if (filePath) {
                        const raw = this.database.importContactFace(
                            filePath[0],
                            name
                        );

                        image.element.src = raw;
                    }
                }]
            });
        }

        this.addElement(faceGallery);

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Generate',
            eventListener: ['click', () => {
                const savePath = dialog.showSaveDialogSync({
                    filters: [
                        { name: 'VCF', extensions: ['vcf'] }
                    ]
                });

                if (savePath) {
                    let vcfOutput = [];

                    const yomiCycle = 'ｱｶｻﾀﾅﾊﾏﾔﾗﾜ';
                    let groupYomi = {};

                    let i = 0;
                    for (const name of groups) {
                        groupYomi[name] = yomiCycle[i % yomiCycle.length];
                        i++;
                    }

                    const contactFaces = this.database.getContactFaces();

                    Object.values(this.database.numbers).forEach(number => {
                        let area;
                        if (number.name in this.database.areas) {
                            area = this.database.areas[number.name];
                        }

                        let noteParts = [];

                        if (area) {
                            noteParts = [
                                `${area.district} District`,
                                `${area.zone} Zone`
                            ];
                        }

                        const encodedNote = encodeQuotedPrintable(noteParts.join(', '));

                        const missionaries = area ? area.people.map(name => {
                            const person = this.database.people[name];
                            return person.type + ' ' + name.split(',')[0];
                        }).join(', ') : '';

                        const group = number.group;

                        const photo = group in contactFaces ? contactFaces[group] : '';

                        const card = generateCard(
                            number.displayName ?? number.name,
                            groupYomi[group],
                            number.number.replace(/^\+81\s*/, ''),
                            missionaries,
                            photo.replace('data:image/png;base64,', '').replace(/^"|"$/g, ''),
                            encodedNote,
                            number.group,
                            number.number.length === 3 ? 'Emergency Contacts' : 'JKM Contacts'
                        );

                        vcfOutput.push(card.replace(/^\s{4}/gm, '').trim());
                    });

                    vcfOutput = vcfOutput.join('\n\n\n');

                    fs.writeFileSync(savePath, vcfOutput);

                    showMessage('Generated contacts file.');
                }
            }]
        });

        this.addElement(button);
    }
}

class CreateGraphsView extends View {
    get name() { return 'Graphs'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Generate Graphs'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'Use this to generate various graphs to be used in MLC, DLC, and for looking at various metrics.<br><br>In order to download these files:<br>1. Access Tableau at this link:<br>https://prod-useast-b.online.tableau.com/t/churchofjesuschrist/views/. If you do not have access, make sure to talk to President Sano!<br>2. Choose the view you want to download data from (Key Indicator Performance or Mission Finding Comparison)<br><br>For Key Indicator Performance:<br>1. Select "Japan Kobe" in the "Mission Name" dropdown<br>2. Ensure the date range is from 1/1/2024 (it will probably default to 1/1/2025) to the current date, or later <br>3. Click the download button in the top right corner and choose the "Crosstab" option<br>4. Ensure the "Missionary KI Table" is the sheet selected, then choose either format of file<br>5. Press Download<br><br>For Mission Finding Comparison:<br>1. Change "Start Date" to 1/1/2001 (some people baptized were found many years ago)<br>2. Ensure "End Date" is the current date, or later<br>3. Click the download button in the top right corner and choose the "Crosstab" option<br>4. Ensure the "Detail" is the sheet selected, then choose either format of file<br>5. Press Download<br><br>You will have to upload these files each time you want to generate new graphs. They will not save. This is to make sure you are using the most up-to-date data from Tableau.'
        });

        this.addElement(comment);

        const fileGallery = new Element('DIV', null, {
            elementClass: 'view-double-gallery'
        });

        const galleryEntryKI = new Element('DIV', fileGallery, {
            elementClass: 'view-gallery-entry'
        });

        new Element('H2', galleryEntryKI, {
            elementClass: 'view-gallery-header',
            text: 'Key Indicator File'
        });

        new Element('IMG', galleryEntryKI, {
            elementClass: 'view-gallery-picture',
            attributes: {
                'SRC': path.join(__dirname, '..', 'assets', 'graphs', 'KITableau.png'),
                'height': '175px',
            }
        });

        const subheaderKI = new Element('H3', galleryEntryKI, {
            elementClass: 'view-gallery-subheader',
            text: 'No file uploaded yet.'
        });
        
        let keyIndicatorFilePath = null;

        new Element('BUTTON', galleryEntryKI, {
            elementClass: 'view-gallery-button',
            text: 'Upload',
            eventListener: ['click', () => {
                const filePath = dialog.showOpenDialogSync({
                    properties: ['openFile'],
                    defaultPath: "C:\\Users\\2016702-REF\\OneDrive - Church of Jesus Christ\\Desktop\\Kobe Desk",
                    filters: [
                        { name: 'KI File', extensions: ['csv', 'xlsx'] }
                    ]
                });
                if (filePath) {
                    keyIndicatorFilePath = filePath[0];
                    subheaderKI.element.textContent = 'Uploaded: ' + keyIndicatorFilePath;
                }
            }]
        });

        const galleryEntryFD = new Element('DIV', fileGallery, {
            elementClass: 'view-gallery-entry'
        });

        new Element('H2', galleryEntryFD, {
            elementClass: 'view-gallery-header',
            text: 'Finding Detail File'
        });

        new Element('IMG', galleryEntryFD, {
            elementClass: 'view-gallery-picture',
            attributes: {
                'SRC': path.join(__dirname, '..', 'assets', 'graphs', 'FDTableau.png'),
                'height': '175px',
            }
        });

        const subheaderFD = new Element('H3', galleryEntryFD, {
            elementClass: 'view-gallery-subheader',
            text: 'No file uploaded yet.'
        });

        let findingDetailFilePath = null;

        new Element('BUTTON', galleryEntryFD, {
            elementClass: 'view-gallery-button',
            text: 'Upload',
            eventListener: ['click', () => {
                const filePath = dialog.showOpenDialogSync({
                    properties: ['openFile'],
                    defaultPath: "C:\\Users\\2016702-REF\\OneDrive - Church of Jesus Christ\\Desktop\\Kobe Desk",
                    filters: [
                        { name: 'FD File', extensions: ['csv', 'xlsx'] }
                    ]
                });

                if (filePath) {
                    findingDetailFilePath = filePath[0];
                    subheaderFD.element.textContent = 'Uploaded: ' + findingDetailFilePath;
                }
            }]
        });

        this.addElement(fileGallery);

        const subheaderOutputPath = new Element('H3', null, {
            elementClass: 'view-gallery-subheader',
            text: 'No output directory selected.'
        });

        const pyDir = path.join(__dirname, '..', 'py');

        let outputFilePath = null;

        const button = new Element('BUTTON', null, {
            elementClass: 'view-button',
            text: 'Select Output Directory',
            eventListener: ['click', () => {

                showMessage('Checking for Python dependencies...');

                const outputPath = dialog.showOpenDialogSync({
                    title: 'Select Output Directory',
                    defaultPath: "C:\\Users\\2016702-REF\\OneDrive - Church of Jesus Christ\\Desktop\\Kobe Desk\\Output Graphs",
                    properties: ['openDirectory']
                });

                if (!outputPath || outputPath.length === 0) {
                    showMessage('Please select an output directory.');
                    return;
                }

                if (outputPath) {
                    outputFilePath = outputPath[0];
                    subheaderOutputPath.element.textContent = 'Output Directory: ' + outputFilePath;
                }

                // Ensure Python dependencies are installed
                const requirementsPath = path.join(pyDir, 'requirements.txt');
                const installProcess = spawn('python', ['-m', 'pip', 'install', '-r', requirementsPath]);

                installProcess.stdout.on('data', (data) => {
                    console.log(`pip stdout: ${data}`);
                });

                installProcess.stderr.on('data', (data) => {
                    console.error(`pip stderr: ${data}`);
                });

                installProcess.on('close', (installCode) => {
                    if (installCode !== 0) {
                        showMessage('Failed to install Python dependencies. See console for details.');
                        return;
                    } else {
                        showMessage('Python dependencies installed successfully.');
                    }
                });
            }]
        });

        this.addElement(subheaderOutputPath);
        this.addElement(button);

        const comment2 = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'Once you have uploaded the files and selected an output directory, you can generate graphs by clicking the "Generate Graph" button for each graph.<br><br>Note: the graphs will generate using last year\'s and this year\'s data.'
        });

        const graphGallery = new Element('DIV', null, {
            elementClass: 'view-double-gallery'
        });

        fs.readdirSync(pyDir).forEach(file => {
            if (file.endsWith('.py')) {
                const graphEntry = new Element('DIV', graphGallery, {
                    elementClass: 'view-gallery-graphEntry'
                });

                new Element('H2', graphEntry, {
                    elementClass: 'view-gallery-header',
                    text: file
                });

                new Element('IMG', graphEntry, {
                    elementClass: 'view-gallery-picture',
                    attributes: {
                        'SRC': path.join(__dirname, '..', 'assets', 'graphs', file.replace('.py', '_en.png')),
                        'height': '175px',
                    }
                });

                new Element('BUTTON', graphEntry, {
                    elementClass: 'view-gallery-button',
                    text: 'Generate Graph',
                    eventListener: ['click', () => {

                        showMessage(`Generating graphs for ${file}...`);

                        if (!keyIndicatorFilePath || !findingDetailFilePath) {
                            showMessage('Please upload both Key Indicator and Finding Detail files before generating graphs.');
                            return;
                        }

                        if (!outputFilePath) {
                            showMessage('Please select an output directory before generating graphs.');
                            return;
                        }

                        const scriptPath = path.join(pyDir, file);
                        const args = [keyIndicatorFilePath, findingDetailFilePath, outputFilePath];
                        const pythonProcess = spawn('python', [scriptPath, ...args], {
                            cwd: pyDir
                        });

                        pythonProcess.stdout.on('data', (data) => {
                            console.log(`${file} stdout: ${data}`);
                        });

                        pythonProcess.stderr.on('data', (data) => {
                            console.error(`${file} stderr: ${data}`);
                        });

                        pythonProcess.on('close', (code) => {
                            if (code === 0) {
                                showMessage(`${file} graphs generated successfully.`);
                            } else {
                                showMessage(`Error generating ${file} graphs. See console for details.`);
                            }
                        });
                    }]
                });
            }
        });

        this.addElement(graphGallery);
        
    }
}

// ...existing code...
class CreateTransferBoardCardsView extends View {
   get name() { return 'Transfer Board Cards'; }

    build() {
        const header = new Element('H1', null, {
            elementClass: 'view-header',
            text: 'Generate Transfer Board Cards'
        });

        this.addElement(header);

        const comment = new Element('DIV', null, {
            elementClass: 'view-comment',
            text: 'Use this to generate printable cards for Transfer Boards.'
        });

        this.addElement(comment);

        // Controls: filter recent arrivals (days), select count/copies
        const controls = new Element('DIV', null, {
            elementClass: 'view-controls'
        });

        new Element('LABEL', controls, {
            elementClass: 'view-label',
            text: 'Filter by MTC within (days):'
        });

        const mtcDaysInput = new Element('INPUT', controls, {
            elementClass: 'view-input-small',
            attributes: { 'type': 'number', 'value': '120', 'min': '0' }
        });

        new Element('BUTTON', controls, {
            elementClass: 'view-button-small',
            text: 'Apply Filter',
            eventListener: ['click', () => {
                const days = parseInt(mtcDaysInput.element.value || '0', 10);
                filterByMtc(days);
            }]
        });

        new Element('LABEL', controls, {
            elementClass: 'view-label',
            text: 'Copies per card:'
        });

        const copiesInput = new Element('INPUT', controls, {
            elementClass: 'view-input-small',
            attributes: { 'type': 'number', 'value': '3', 'min': '1' }
        });

        this.addElement(controls);

        // Selection list
        const selectionContainer = new Element('DIV', null, {
            elementClass: 'view-multi-select'
        });

        // Gather people + profiles safely (fallbacks)
        const peopleMap = this.database.people || {};
        const profiles = (typeof this.database.getProfiles === 'function') ? this.database.getProfiles() : (this.database.profiles || {});

        const personEntries = [];

        // Build list UI
        for (const name of Object.keys(peopleMap)) {
            const person = peopleMap[name] || {};

            const entry = new Element('DIV', selectionContainer, {
                elementClass: 'view-select-entry'
            });

            const cb = new Element('INPUT', entry, {
                elementClass: 'view-select-checkbox',
                attributes: { 'type': 'checkbox', 'value': name }
            });

            const thumb = new Element('IMG', entry, {
                elementClass: 'view-select-thumb',
                attributes: {
                    'SRC': profiles[name] || '',
                    'alt': name
                }
            });

            const labelWrap = new Element('DIV', entry, {
                elementClass: 'view-select-labels'
            });

            new Element('DIV', labelWrap, {
                elementClass: 'view-select-name',
                text: name
            });

            // Try a few common keys for MTC/hometown
            const mtc = person.mtcStart || person.mtc_start || person.mtc || person.arrivalDate || '';
            const hometown = person.hometown || person.city || person.home || '';

            new Element('DIV', labelWrap, {
                elementClass: 'view-select-sub',
                text: `MTC: ${mtc || 'N/A'}  •  Hometown: ${hometown || 'N/A'}`
            });

            personEntries.push({
                name,
                person,
                entryElement: entry,
                checkbox: cb,
                mtc,
                hometown,
                photo: profiles[name] || ''
            });
        }

        this.addElement(selectionContainer);

        // Helper to parse date (very forgiving)
        function parseDateSafe(s) {
            if (!s) return null;
            // try ISO or YYYY/MM/DD or MM/DD/YYYY
            const d = new Date(s);
            if (!isNaN(d.getTime())) return d;
            // try replace - with /
            const d2 = new Date(s.replace(/-/g, '/'));
            if (!isNaN(d2.getTime())) return d2;
            return null;
        }

        function filterByMtc(days) {
            const now = new Date();
            for (const e of personEntries) {
                if (!e.mtc) {
                    e.entryElement.element.style.display = '';
                    continue;
                }
                const dt = parseDateSafe(e.mtc);
                if (!dt) {
                    e.entryElement.element.style.display = '';
                    continue;
                }
                const diffDays = Math.floor((now - dt) / (1000 * 60 * 60 * 24));
                e.entryElement.element.style.display = (diffDays <= days) ? '' : 'none';
            }
        }

        // Buttons: Preview and Generate
        const actions = new Element('DIV', null, {
            elementClass: 'view-actions'
        });

        new Element('BUTTON', actions, {
            elementClass: 'view-button',
            text: 'Preview',
            eventListener: ['click', () => {
                const selected = personEntries.filter(p => p.checkbox.element.checked && p.entryElement.element.style.display !== 'none');
                if (!selected.length) {
                    showMessage('No missionaries selected for preview.');
                    return;
                }

                const cards = selected.map(p => ({
                    name: p.name,
                    mtc: p.mtc || '',
                    hometown: p.hometown || '',
                    photo: p.photo || ''
                }));

                const payload = {
                    cards,
                    copies: parseInt(copiesInput.element.value || '3', 10),
                    cardSizeInches: { width: 3, height: 5 }, // per your spec
                    paperSize: 'A4',
                    preview: true
                };

                // Open generate window visible
                ipcRenderer.sendSync('create-window', 'generate', {
                    show: true,
                    webPreferences: {
                        nodeIntegration: true,
                        enableRemoteModule: true,
                        contextIsolation: false
                    },
                    titleBarStyle: 'hidden',
                    titleBarOverlay: { color: '#292929', symbolColor: '#fff' },
                    minWidth: 600,
                    minHeight: 450,
                    backgroundColor: '#fff'
                });

                // Send data to main/process or generate window handler
                ipcRenderer.send('transfer-board-cards-data', payload);
            }]
        });

        new Element('BUTTON', actions, {
            elementClass: 'view-button',
            text: 'Generate PDF',
            eventListener: ['click', () => {
                const selected = personEntries.filter(p => p.checkbox.element.checked && p.entryElement.element.style.display !== 'none');
                if (!selected.length) {
                    showMessage('No missionaries selected for generation.');
                    return;
                }

                const savePath = dialog.showSaveDialogSync({
                    filters: [{ name: 'PDF', extensions: ['pdf'] }],
                    defaultPath: path.join(require('os').homedir(), 'transfer_board_cards.pdf')
                });

                if (!savePath) return;

                const cards = selected.map(p => ({
                    name: p.name,
                    mtc: p.mtc || '',
                    hometown: p.hometown || '',
                    photo: p.photo || ''
                }));

                const payload = {
                    cards,
                    copies: parseInt(copiesInput.element.value || '3', 10),
                    cardSizeInches: { width: 3, height: 5 },
                    paperSize: 'A4',
                    preview: false,
                    savePath
                };

                // Prefer creating a hidden generate window to render and save PDF
                ipcRenderer.sendSync('create-window', 'generate', {
                    show: false,
                    webPreferences: {
                        nodeIntegration: true,
                        enableRemoteModule: true,
                        contextIsolation: false
                    },
                    titleBarStyle: 'hidden',
                    titleBarOverlay: { color: '#292929', symbolColor: '#fff' },
                    minWidth: 600,
                    minHeight: 450,
                    backgroundColor: '#fff'
                });

                ipcRenderer.send('transfer-board-cards-data', payload);
            }]
        });

        this.addElement(actions);
    }
}
// ...existing code...