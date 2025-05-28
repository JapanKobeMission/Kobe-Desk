const { ipcRenderer } = require("electron");

class Database {
    constructor() {
        this.areas = {};
        this.people = {};
        this.numbers = {};
        this.addresses = {};
        this.teams = {};
    }

    addArea(area) {
        this.areas[area.name] = area;
    }

    addPerson(person) {
        this.people[person.name] = person;
    }

    addNumber(number) {
        this.numbers[number.number] = number;
    }

    updateNumber(number, area, group, displayName) {
        number = PhoneNumber.normalize(number);

        if (number in this.numbers) {
            this.numbers[number].name = area;
            this.numbers[number].group = group;
            this.numbers[number].displayName = displayName ?? this.numbers[number].displayName;
        } else {
            this.numbers[number] = new PhoneNumber(
                number,
                area,
                group,
                displayName,
                null
            );
        }
    }

    addAddress(address) {
        this.addresses[address.name] = address;
    }

    addTeam(team) {
        this.teams[team.name] = team;
    }

    importTable(table) {
        this.areas = {};
        this.people = {};

        for (const row of table) {
            if (row['Phone']) {
                for (const phone of row['Phone'].split(', ')) {
                    this.updateNumber(phone, row['Area'], row['Zone']);
                }
            }

            if (row['Area']) {
                if (row['Area'] in this.areas) {
                    this.areas[row['Area']].addPerson(
                        row['Missionary'].trim()
                    );
                } else {
                    const area = new Area(
                        row['Area'],
                        row['District'],
                        row['Zone'],
                        [row['Missionary'].trim()]
                    );
                    this.addArea(area);
                }
            }

            if (row['Missionary']) {
                const person = new Person(
                    row['Missionary'].trim(),
                    row['ID'],
                    row['Type'],
                    row['Position Abbr'],
                    row['Status'],
                    row['Arrival Date'],
                    row['Release Date']
                );
                this.addPerson(person);
            }
        }

        this.saveData();
        
        showMessage('Imported table.');
    }

    importAddresses(table) {
        for (const row of table) {
            const fields = ['Name', 'Postal Code', 'English Address', 'Japanese Address'];

            if (fields.every(x => x in row && row[x])) {
                this.addAddress(new Address(
                    row['Name'],
                    row['Postal Code'],
                    row['English Address'],
                    row['Japanese Address'],
                    []
                ));
            }
        }

        this.saveData();
        
        showMessage('Imported addresses.');
    }

    importAddressKey(table) {
        for (const row of table) {
            const fields = ['Areas', 'Apartment'];

            if (fields.every(x => x in row && row[x])) {
                const area = row['Areas'];
                const apartmentName = row['Apartment'];

                if (area in this.areas && apartmentName in this.addresses) {
                    this.addresses[apartmentName].areas.push(area);
                }
            }
        }

        this.saveData();
        
        showMessage('Imported address key.');
    }

    importProfiles(data) {
        for (const ID in data) {
            const { extension, filePath } = data[ID];

            ipcRenderer.sendSync(
                'save-image',
                'Profiles',
                `${ID}.jpg`,
                filePath,
                extension === 'txt',
                281,
                337
            );
        }
        
        showMessage('Imported profiles.');
    }

    importCovers(data) {
        for (const ID in data) {
            const { extension, filePath } = data[ID];

            ipcRenderer.sendSync(
                'save-image',
                'Covers',
                `${ID}.png`,
                filePath,
                extension === 'txt',
                600,
                849
            );
        }
        
        showMessage('Imported covers.');
    }

    importCover(filePath, zoneName, freeSize) {
        const cover = ipcRenderer.sendSync(
            'save-image',
            'Covers',
            `${zoneName}.png`,
            filePath,
            false,
            freeSize ? undefined : 600,
            freeSize ? undefined : 849
        );
        
        showMessage('Imported cover.');

        return cover;
    }

    importProfile(filePath, ID) {
        const profile = ipcRenderer.sendSync(
            'save-image',
            'Profiles',
            `${ID}.jpg`,
            filePath,
            false,
            281,
            337
        );
        
        showMessage('Imported profile.');

        return profile;
    }

    importContactFace(filePath, zoneName) {
        const contactFace = ipcRenderer.sendSync(
            'save-image',
            'ContactFaces',
            `${zoneName}.png`,
            filePath,
            false,
            300,
            300
        );
        
        showMessage('Imported contact face.');

        return contactFace;
    }

    saveData() {
        ipcRenderer.sendSync('save-data', [
            this.areas,
            this.people,
            this.numbers,
            this.addresses,
            this.teams
        ]);
    }

    loadData() {
        const data = ipcRenderer.sendSync('load-data');

        if (Array.isArray(data) && data.length === 5) {
            const [areas, people, numbers, addresses, teams] = data;

            for (const area of Object.values(areas)) {
                this.addArea(
                    new Area(
                        area.name,
                        area.district,
                        area.zone,
                        area.people.map(person => person.trim())
                    )
                );
            }

            for (const person of Object.values(people)) {
                this.addPerson(
                    new Person(
                        person.name.trim(),
                        person.ID,
                        person.type,
                        person.assignment,
                        person.status,
                        person.arrivalDate,
                        person.releaseDate
                    )
                );
            }

            for (const number of Object.values(numbers)) {
                this.addNumber(
                    new PhoneNumber(
                        number.number,
                        number.name,
                        number.group,
                        number.displayName,
                        number.lineAssignment
                    )
                );
            }

            for (const address of Object.values(addresses)) {
                this.addAddress(
                    new Address(
                        address.name,
                        address.postalCode,
                        address.englishAddress,
                        address.japaneseAddress,
                        address.areas
                    )
                );
            }

            for (const team of Object.values(teams)) {
                this.addTeam(
                    new Team(
                        team.name,
                        team.people,
                        team.roles
                    )
                );
            }
        }
    }

    getZoneAreas() {
        let zones = {};

        for (const areaName in this.areas) {
            const area = this.areas[areaName];

            if (area.zone in zones) {
                zones[area.zone].push(area.name);
            } else {
                zones[area.zone] = [area.name];
            }
        }

        return zones;
    }

    getZoneDistricts() {
        let zones = {};

        for (const areaName in this.areas) {
            const area = this.areas[areaName];

            if (area.zone in zones) {
                if (area.district in zones[area.zone]) {
                    zones[area.zone][area.district].push(area.name);
                } else {
                    zones[area.zone][area.district] = [area.name];
                }
            } else {
                zones[area.zone] = {};
                zones[area.zone][area.district] = [area.name];
            }
        }

        return zones;
    }

    getAreaNumber(area) {
        return Object.values(this.numbers).find(number => {
            return number.name === area;
        });
    }

    getPersonArea(person) {
        return Object.values(this.areas).find(area => {
            return area.people.indexOf(person) > -1
        });
    }

    getAllAreaNumbers(area) {
        return Object.values(this.numbers).filter(number => {
            return number.name === area;
        });
    }

    getAreaAddress(area) {
        return Object.values(this.addresses).find(address => {
            return address.areas.indexOf(area) > -1;
        });
    }

    getProfiles() {
        const data = ipcRenderer.sendSync('get-images', 'Profiles');

        let profiles = {};

        for (const image of data) {
            profiles[image.basename] = image.raw;
        }

        return profiles;
    }

    getCovers() {
        const data = ipcRenderer.sendSync('get-images', 'Covers');

        let covers = {};

        for (const image of data) {
            covers[image.basename] = image.raw;
        }

        return covers;
    }

    getContactFaces() {
        const data = ipcRenderer.sendSync('get-images', 'ContactFaces');

        let contactFaces = {};

        for (const image of data) {
            contactFaces[image.basename] = image.raw;
        }

        return contactFaces;
    }
}

class Area {
    constructor(name, district, zone, people) {
        this.name = name;
        this.district = district;
        this.zone = zone;
        this.people = people;
    }

    addPerson(name) {
        this.people.push(name.trim());
    }
}

class Person {
    constructor(name, ID, type, assignment, status, arrivalDate, releaseDate) {
        this.name = name;
        this.ID = ID;
        this.type = type;
        this.assignment = assignment;
        this.status = status;
        this.arrivalDate = arrivalDate;
        this.releaseDate = releaseDate;
    }
}

class PhoneNumber {
    constructor(number, name, group, displayName, lineAssignment) {
        this.number = PhoneNumber.normalize(number);
        this.name = name;
        this.group = group;
        this.displayName = displayName
        if (!this.displayName) {
            this.displayName = this.name;
        }
        this.lineAssignment = lineAssignment;
    }

    static normalize(number) {
        const plain = number.replace(/\D/g, '');

        if (plain.length <= 3) {
            return number;
        }

        return `+${plain.slice(0, 2)} ${plain.slice(2, 5)}-${plain.slice(5, 9)}-${plain.slice(9)}`;
    }
}

class Address {
    constructor(name, postalCode, englishAddress, japaneseAddress, areas) {
        this.name = name;
        this.postalCode = postalCode;
        this.englishAddress = englishAddress;
        this.japaneseAddress = japaneseAddress;
        this.areas = areas;
    }
}

class Team {
    constructor(name, people, roles) {
        this.name = name;
        this.people = people;
        this.roles = roles;
    }
}