const database = new Database();
database.loadData();

const contentElement = new Element('DIV', document.body, {
    ID: 'content'
});

const covers = database.getCovers();

const titleCoverWrapper = new Element('DIV', contentElement, {
    elementClass: 'cover-wrapper'
});

new Element('IMG', titleCoverWrapper, {
    elementClass: 'cover',
    attributes: {
        src: covers['Cover']
    }
});

new Element('H1', contentElement, {
    elementClass: 'section-title',
    text: 'Phone List'
});

const phoneListSection = new Element('DIV', contentElement, {
    elementClass: 'section',
    ID: 'phone-list'
});

function sortObject(unordered) {
    return Object.keys(unordered).sort().reduce(
        (obj, key) => {
            obj[key] = unordered[key];
            return obj;
        },
        {}
    );
}

let phoneIndex = {};

for (const number of Object.values(database.numbers)) {
    if (number.group in phoneIndex) {
        phoneIndex[number.group].push(number);
    } else {
        phoneIndex[number.group] = [number];
    }
}

phoneIndex = sortObject(phoneIndex);

for (const group in phoneIndex) {
    if (group === "Extra") {
        continue;
    }
    const table = new Element('TABLE', phoneListSection, {
        elementClass: 'phone-table'
    });

    const headerRow = new Element('TR', table, {
        elementClass: 'phone-table-header-row'
    });

    new Element('TH', headerRow, {
        elementClass: 'phone-table-header',
        text: group,
        attributes: {
            colspan: 2
        }
    });

    for (const number of phoneIndex[group]) {
        if (group === "Extra") {
            continue;
        }
        const row = new Element('TR', table, {
            elementClass: 'phone-table-row'
        });

        new Element('TD', row, {
            elementClass: 'phone-table-row-name',
            text: number.name
        });

        new Element('TD', row, {
            elementClass: 'phone-table-row-number',
            text: number.number
        });
    }
}

new Element('H1', contentElement, {
    elementClass: 'section-title',
    text: 'Address List'
});

const addressListSection = new Element('DIV', contentElement, {
    elementClass: 'section',
    ID: 'address-list'
});

const zoneAreas = sortObject(database.getZoneAreas());

for (const zone in zoneAreas) {
    new Element('H2', addressListSection, {
        elementClass: 'address-zone-header',
        text: zone
    });

    for (const areaName of zoneAreas[zone]) {
        const address = database.getAreaAddress(areaName);

        if (!address) {
            ipcRenderer.sendSync(
                'append-warning',
                `There is no address listed for ${areaName}.`
            );
            continue;
        }

        const table = new Element('TABLE', addressListSection, {
            elementClass: 'address-table'
        });

        const headerRow = new Element('TR', table, {
            elementClass: 'address-table-header-row'
        });

        new Element('TH', headerRow, {
            elementClass: 'address-table-header',
            text: areaName,
            attributes: {
                colspan: 2
            }
        });

        const topRow = new Element('TR', table, {
            elementClass: 'address-table-row'
        });

        const bottomRow = new Element('TR', table, {
            elementClass: 'address-table-row'
        });

        new Element('TD', topRow, {
            elementClass: 'address-table-row-postal',
            text: address.postalCode,
            attributes: {
                rowspan: 2
            }
        });

        new Element('TD', topRow, {
            elementClass: 'address-table-row-english',
            text: address.englishAddress
        });

        new Element('TD', bottomRow, {
            elementClass: 'address-table-row-japanese',
            text: address.japaneseAddress
        });
    }
}

const profiles = database.getProfiles();

const zoneDistricts = database.getZoneDistricts();

for (const zone in zoneDistricts) {
    const coverWrapper = new Element('DIV', contentElement, {
        elementClass: 'cover-wrapper'
    });

    new Element('IMG', coverWrapper, {
        elementClass: 'cover',
        attributes: {
            src: covers[zone]
        }
    });

    for (const district in zoneDistricts[zone]) {
        const districtSection = new Element('DIV', contentElement, {
            elementClass: 'district-section'
        });

        new Element('H2', districtSection, {
            elementClass: 'district-header',
            text: district
        });

        let i = 0;
        let row;

        for (const areaName of zoneDistricts[zone][district]) {
            const area = database.areas[areaName];

            if (i % 2 === 0) {
                row = new Element('DIV', districtSection, {
                    elementClass: 'district-row'
                });
            }

            const table = new Element('TABLE', row, {
                elementClass: 'area-table'
            });

            const headerRow = new Element('TR', table, {
                elementClass: 'area-header-row'
            });

            new Element('TH', headerRow, {
                elementClass: 'area-header',
                text: areaName,
                attributes: {
                    colspan: area.people.length
                }
            });

            const profileRow = new Element('TR', table, {
                elementClass: 'area-row'
            });
            const nameRow = new Element('TR', table, {
                elementClass: 'name-row'
            });
            const roleRow = new Element('TR', table, {
                elementClass: 'role-row'
            });

            for (const personName of area.people) {
                const person = database.people[personName];

                const imageRaw = profiles[person.ID];

                const imageColumn = new Element('TD', profileRow, {
                    elementClass: 'profile-column'
                });

                new Element('IMG', imageColumn, {
                    elementClass: 'profile-image',
                    attributes: {
                        src: imageRaw
                    }
                });

                new Element('TD', nameRow, {
                    elementClass: 'name-column',
                    text: person.type + ' ' + person.name
                });

                new Element('TD', roleRow, {
                    elementClass: 'role-column',
                    text: person.role
                });
            }

            for (const number of database.getAllAreaNumbers(area.name)) {
                const row = new Element('TR', table, {
                    elementClass: 'area-number-row'
                });

                new Element('TD', row, {
                    elementClass: 'area-number-column',
                    text: number.number,
                    attributes: {
                        colspan: area.people.length
                    }
                });
            }

            const address = database.getAreaAddress(areaName);

            if (address) {
                const japaneseAddressRow = new Element('TR', table, {
                    elementClass: 'area-address-row'
                });

                new Element('TD', japaneseAddressRow, {
                    elementClass: 'japanese-address-column',
                    text: address.japaneseAddress,
                    attributes: {
                        colspan: area.people.length
                    }
                });

                const englishAddressRow = new Element('TR', table, {
                    elementClass: 'area-address-row'
                });

                new Element('TD', englishAddressRow, {
                    elementClass: 'english-address-column',
                    text: address.englishAddress,
                    attributes: {
                        colspan: area.people.length
                    }
                });
            }

            i++;
        }

        if (i % 2 !== 0) {
            new Element('DIV', row, {
                elementClass: 'table-filler'
            });
        }
    }
}

new Element('H1', contentElement, {
    elementClass: 'section-title',
    text: 'Teams'
});

const teamsSection = new Element('DIV', contentElement, {
    elementClass: 'section',
    ID: 'teams'
});

for (const team of Object.values(database.teams)) {
    new Element('TR', teamsSection, {
        elementClass: 'team-list-header',
        text: team.name
    });

    const list = new Element('UL', teamsSection, {
        elementClass: 'team-list'
    });

    for (const personName of team.people) {
        new Element('LI', list, {
            elementClass: [
            'team-list-element',
            `team-list-element-${team.roles[personName]}`
            ],
            text: database.people[personName].type + ' ' + personName
        });

        
    }
}

new Element('H1', contentElement, {
    elementClass: 'section-title',
    text: 'Transfer Board'
});

const transferBoardWrapper = new Element('DIV', contentElement, {
    elementClass: 'nonbreak-cover-wrapper'
});

new Element('IMG', transferBoardWrapper, {
    elementClass: 'cover',
    attributes: {
        src: covers['Transfer Board']
    }
});

new Element('H1', contentElement, {
    elementClass: 'section-title',
    text: 'Calendar'
});

const calendarWrapper = new Element('DIV', contentElement, {
    elementClass: 'nonbreak-cover-wrapper'
});

new Element('IMG', calendarWrapper, {
    elementClass: 'cover',
    attributes: {
        src: covers['Calendar']
    }
});

new Element('BUTTON', contentElement, {
    elementClass: 'save-button',
    text: 'SAVE',
    eventListener: ['click', () => {
        ipcRenderer.sendSync('save-pdf');
    }]
});

contentElement.render();