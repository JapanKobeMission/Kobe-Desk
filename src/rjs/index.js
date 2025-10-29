const database = new Database();
database.loadData();

const viewCategories = {
    'Import': [
        new ImportRosterView(database),
        new ImportAddressesView(database),
        new ImportAddressKeyView(database),
        new ImportProfilesView(database),
        new ImportCoversView(database)
    ],
    'Create': [
        new CreateTransferDocsView(database),
        new CreateContactsView(database),
        new CreateGraphsView(database),
        new CreateTransferBoardCardsView(database)
    ],
    'Edit': [
        new EditPeopleView(database),
        new EditAddressesView(database),
        new EditContactsView(database),
        new EditTeamsView(database)
    ]
};

const viewNavigator = new ViewNavigator();

for (const category in viewCategories) {
    for (const view of viewCategories[category]) {
        viewNavigator.addView(category, view);
    }
}

viewNavigator.selectCategory('Edit');
viewNavigator.renderTitlebar();