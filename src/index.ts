import g from '@battis/gas-lighter';

type FolderConfig = {
  email: string;
  attach?: boolean;
  separate?: boolean;
};

const TRIGGER = 'trigger';
const LAST_CHECK = 'lastCheck';

function getFolderIds() {
  return PropertiesService.getUserProperties()
    .getKeys()
    .filter((key) => key !== TRIGGER && key !== LAST_CHECK);
}

global.onHomepage = (e) => {
  const stats = [];
  const hasTrigger = !!g.PropertiesService.getUserProperty(TRIGGER);
  stats.push(`Hourly check is${hasTrigger ? '' : ' not'} active`);
  if (hasTrigger) {
    stats.push(
      `Last checked: ${new Date(
        g.PropertiesService.getUserProperty(LAST_CHECK)
      ).toLocaleString()}`
    );
    stats.push(
      g.CardService.Widget.newTextButton({
        text: 'Check now',
        functionName: 'hourly'
      })
    );
  } else {
    g.PropertiesService.deleteUserProperty(LAST_CHECK);
  }

  const title = CardService.newCardSection().addWidget(
    CardService.newImage().setImageUrl(
      'https://groton-school.github.io/email-new-drive-files/store/card-banner-assets/cardservice-header.png'
    )
  );

  const folders = getFolderIds().map((folder) =>
    g.CardService.Widget.newTextButton({
      text: DriveApp.getFolderById(folder).getName(),
      functionName: 'view',
      parameters: { folder }
    })
  );

  let foldersSection =
    CardService.newCardSection().setHeader('Monitored folders');
  folders.forEach(
    (widget) => (foldersSection = foldersSection.addWidget(widget))
  );
  if (!folders.length) {
    foldersSection.addWidget(g.CardService.Widget.newTextParagraph('None'));
  }

  let statsSection = CardService.newCardSection().setHeader('Activity');
  stats.forEach(
    (item) =>
      (statsSection = statsSection.addWidget(
        typeof item === 'string'
          ? g.CardService.Widget.newTextParagraph(item)
          : item
      ))
  );

  return g.CardService.Card.create({
    widgets: [, 'Select a Drive folder at left to monitor for new files.'],
    sections: [title, foldersSection, statsSection]
  });
};

global.view = (e) => {
  const folder = DriveApp.getFolderById(e.parameters.folder);
  return folderList([{ id: folder.getId(), title: folder.getName() }]);
};

global.onItemsSelected = (e) => {
  const folders = e.drive.selectedItems.filter(
    (item) => item.mimeType === 'application/vnd.google-apps.folder' // TODO MimeType.FOLDER won't compile
  );
  if (folders.length) {
    return folderList(folders);
  }
  return global.home();
};

function folderList(folders) {
  return g.CardService.Card.create({
    sections: folders.map((folder) => {
      const saved = g.PropertiesService.getUserProperty(folder.id);
      const section = CardService.newCardSection()
        .setHeader(folder.title)
        .addWidget(
          CardService.newTextInput()
            .setFieldName(`${folder.id}/email`)
            .setValue(saved ? saved.email : '')
            .setTitle('Email recipient')
            .setHint('user@example.com')
        )
        .addWidget(
          CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setFieldName(`${folder.id}/attach`)
            .addItem(
              'Include file(s) as attachment',
              true,
              saved ? !!saved.attach : false
            )
        )
        .addWidget(
          CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setFieldName(`${folder.id}/separate`)
            .addItem(
              'Separate notification for each file',
              true,
              saved ? !!saved.separate : false
            )
        );
      if (saved) {
        section.addWidget(
          g.CardService.Widget.newTextButton({
            text: 'Remove',
            functionName: 'remove',
            parameters: { folder: folder.id }
          })
        );
      }
      return section;
    }),
    widgets: [
      g.CardService.Widget.newTextButton({
        text: 'Save',
        functionName: 'save'
      }).setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    ]
  });
}

global.home = () => {
  return CardService.newNavigation()
    .popToRoot()
    .updateCard(global.onHomepage());
};

global.remove = (e) => {
  g.PropertiesService.deleteUserProperty(e.parameters.folder);
  return g.CardService.Card.create({
    header: 'Removed',
    widgets: [
      g.CardService.Widget.newTextButton({ text: 'Ok', functionName: 'home' })
    ]
  });
};

global.save = (e) => {
  const folders = {};
  Object.keys(e.commonEventObject.formInputs).forEach((key) => {
    const [id, field] = key.split('/');
    if (!folders[id]) {
      folders[id] = {};
    }
    const value = e.commonEventObject.formInputs[key].stringInputs.value.join();
    folders[id][field] = field === 'email' ? value : true;
  });
  for (const id in folders) {
    g.PropertiesService.setUserProperty(id, folders[id]);
  }

  if (!g.PropertiesService.getUserProperty(TRIGGER)) {
    g.PropertiesService.setUserProperty(LAST_CHECK, new Date().toISOString());
    const trigger = ScriptApp.newTrigger('hourly')
      .timeBased()
      .everyHours(1)
      .create();
    g.PropertiesService.setUserProperty(TRIGGER, trigger.getUniqueId());
  }

  return g.CardService.Card.create({
    header: 'Saved',
    widgets: [
      g.CardService.Widget.newTextButton({ text: 'Ok', functionName: 'home' })
    ]
  });
};

function safeSend(email, subject, body, attachments) {
  try {
    GmailApp.sendEmail(email, subject, body, { attachments });
  } catch (error) {
    Logger.log({
      message: `failed to send ${email}: ${subject}`,
      body,
      attachments,
      error
    });
  }
}

global.hourly = () => {
  const currentCheck = new Date().toISOString();
  const lastCheck = new Date(g.PropertiesService.getUserProperty(LAST_CHECK));
  for (const id of getFolderIds()) {
    const folder = DriveApp.getFolderById(id);
    const config = g.PropertiesService.getUserProperty(id) as FolderConfig;
    const files = folder.getFiles(); // DriveApp is Drive v2, not v3
    const summaries = [];
    let attachments = config.attach ? [] : undefined;
    while (files.hasNext()) {
      const file = files.next();
      if (file.getDateCreated() >= lastCheck) {
        let summary = `${file.getName()}: ${file.getUrl()}`;
        if (attachments !== undefined) {
          try {
            attachments.push(file.getAs('application/pdf'));
          } catch (e) {
            attachments.push(file.getBlob()); // fall back to native type
          }
        }
        if (config.separate) {
          safeSend(
            config.email,
            `${file.getName()} added to ${folder.getName()}`,
            summary,
            attachments
          );
          if (attachments) {
            attachments = [];
          }
        } else {
          summaries.push(summary);
        }
      }
      if (summaries.length) {
        safeSend(
          config.email,
          `Files added to ${folder.getName}`,
          summaries.join('\n\n'),
          attachments
        );
      }
    }
  }
  g.PropertiesService.setUserProperty(LAST_CHECK, currentCheck);
};
