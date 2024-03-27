import pkg from '../package.json';
import g from '@battis/gas-lighter';

type FolderConfig = {
  email: string;
  attach?: boolean;
  separate?: boolean;
};

type FileSummary = {
  name: string;
  id: string;
  sent: boolean;
};

type FolderSummary = {
  name: string;
  id: string;
  files: FileSummary[];
};

type CheckSummary = {
  folders: FolderSummary[];
  attempted: number;
  sent: number;
};

const TRIGGER = 'trigger';
const TRIGGER_VERSION = 'triggerVersion';
const LAST_CHECK = 'lastCheck';
const LAST_CHECK_SUMMARY = 'lastCheckSummary';
const PROPERTIES = [TRIGGER, TRIGGER_VERSION, LAST_CHECK, LAST_CHECK_SUMMARY];

function getFolderIds() {
  return PropertiesService.getUserProperties()
    .getKeys()
    .filter((key) => PROPERTIES.indexOf(key) === -1);
}

global.onHomepage = () => {
  const stats = [];
  const hasTrigger = !!g.PropertiesService.getUserProperty(TRIGGER);
  let triggerVersion = g.PropertiesService.getUserProperty(TRIGGER_VERSION);
  if (triggerVersion !== pkg.version) {
    triggerVersion = resetTriggerToCurrentAddOnVersion();
  }
  stats.push(
    `Hourly check is${hasTrigger ? '' : ' not'} active (${triggerVersion})`
  );
  if (hasTrigger) {
    stats.push(
      `Last checked: ${new Date(
        g.PropertiesService.getUserProperty(LAST_CHECK)
      ).toLocaleString()}`
    );
    stats.push(...displaySummary());
    stats.push(
      g.CardService.Widget.newTextButton({
        text: 'Check now',
        functionName: 'checkNow'
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
  resetTriggerToCurrentAddOnVersion();
  return g.CardService.Card.create({
    header: 'Saved',
    widgets: [
      g.CardService.Widget.newTextButton({ text: 'Ok', functionName: 'home' })
    ]
  });
};

function displaySummary(): string[] {
  const summary: CheckSummary =
    g.PropertiesService.getUserProperty(LAST_CHECK_SUMMARY);
  return summary.folders.map(
    (folder) =>
      `${folder.files.length} new files in ${folder.name}${
        folder.files.length > 0
          ? ` (${folder.files
              .map(
                (file) =>
                  `<a href="https://drive.google.com/file/d/${file.id}">${file.name}</a>`
              )
              .join(', ')})`
          : ''
      }.`
  );
}

global.checkNow = () => {
  global.hourly();
  const lastCheck = new Date(
    g.PropertiesService.getUserProperty(LAST_CHECK)
  ).toLocaleString();
  return g.CardService.Card.create({
    header: `Checked at ${lastCheck}.`,
    widgets: [
      ...displaySummary(),
      g.CardService.Widget.newTextButton({ text: 'Ok', functionName: 'home' })
    ]
  });
};

function resetTriggerToCurrentAddOnVersion() {
  // delete old trigger
  const triggerid = g.PropertiesService.getUserProperty(TRIGGER);
  if (triggerid) {
    ScriptApp.getProjectTriggers().forEach((trigger) => {
      if (trigger.getUniqueId() === triggerid) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }
  global.hourly();
  // create a new trigger
  const trigger = ScriptApp.newTrigger('hourly')
    .timeBased()
    .everyHours(1)
    .create();
  g.PropertiesService.setUserProperty(TRIGGER, trigger.getUniqueId());
  g.PropertiesService.setUserProperty(TRIGGER_VERSION, pkg.version);
  return pkg.version;
}

function safeSend(
  email: string,
  subject: string,
  body: string,
  attachments: GoogleAppsScript.Base.BlobSource[]
) {
  try {
    GmailApp.sendEmail(email, subject, body, { attachments });
    return true;
  } catch (error) {
    Logger.log({
      message: `failed to send ${email}: ${subject}`,
      body,
      attachments,
      error
    });
    return false;
  }
}

global.hourly = () => {
  const currentCheck = new Date().toISOString();
  const lastCheck = new Date(g.PropertiesService.getUserProperty(LAST_CHECK));
  const lastCheckSummary: CheckSummary = {
    folders: [],
    attempted: 0,
    sent: 0
  };
  for (const id of getFolderIds()) {
    const folder = DriveApp.getFolderById(id);
    const folderSummary: FolderSummary = {
      name: folder.getName(),
      id: folder.getId(),
      files: []
    };
    const config = g.PropertiesService.getUserProperty(id) as FolderConfig;
    const files = folder.getFiles(); // DriveApp is Drive v2, not v3
    const summaries = [];
    let attachments = config.attach ? [] : undefined;
    while (files.hasNext()) {
      const file = files.next();
      const fileSummary: FileSummary = {
        name: file.getName(),
        id: file.getId(),
        sent: false
      };
      if (file.getDateCreated() >= lastCheck) {
        let summary = `${file.getName()}: ${file.getUrl()}`;
        if (attachments !== undefined) {
          try {
            attachments.push(file.getAs('application/pdf'));
          } catch (e) {
            attachments.push(file.getBlob()); // fall back to native type
          }
        }
        if ((fileSummary.sent = config.separate)) {
          lastCheckSummary.attempted++;
          if (
            safeSend(
              config.email,
              `${file.getName()} added to ${folder.getName()}`,
              summary,
              attachments
            )
          ) {
            lastCheckSummary.sent++;
          }
          if (attachments) {
            attachments = [];
          }
        } else {
          summaries.push(summary);
        }
        folderSummary.files.push(fileSummary);
      }
    }
    if (summaries.length) {
      lastCheckSummary.attempted++;
      if (
        safeSend(
          config.email,
          `Files added to ${folder.getName()}`,
          summaries.join('\n\n'),
          attachments
        )
      ) {
        folderSummary.files.forEach((f) => (f.sent = true));
        lastCheckSummary.sent++;
      }
    }
    lastCheckSummary.folders.push(folderSummary);
  }
  g.PropertiesService.setUserProperty(LAST_CHECK, currentCheck);
  g.PropertiesService.setUserProperty(LAST_CHECK_SUMMARY, lastCheckSummary);
  Logger.log(lastCheckSummary);
};
