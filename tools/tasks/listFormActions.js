//@ts-check

const { customTask } = require('sp-build-tasks');
// const { create } = require('sp-request');
const { JsomNode } = require('sp-jsom-node');
const { getConfigs } = require('sp-build-tasks/dist/tasks/config');
const { setupPnp } = require('sp-build-tasks/dist/utils/pnpjs');
const inquirer = require('inquirer');
// const { sp, Web, WebPartsPersonalizationScope } = require('@pnp/sp');


const handleError = (e, cb) => {
  console.log(e);
  cb();
};

const promptUser = async (lists) => {
  let excludeLists = [
    'Content and Structure Reports',
    'Form Templates',
    'MicroFeed',
    'Reusable Content',
    'Site Collection Documents',
    'Site Collection Images',
    'Site Pages',
    'Pages',
    'Workflow Tasks'
    ];

  let questions = [
    {
      type: 'list',
      name: 'list',
      message: 'Please select a list',
      choices: lists.reduce((prev, l) => {
        if (!excludeLists.includes(l.title)) {
          prev.push(l.title);
        }
        return prev;
      }, [])
    },
    {
      type: 'list',
      name: 'forms',
      message: 'Which form should be replaced?',
      choices: ['New', 'Edit', 'Display', 'All']
    }
  ];

  let answers = await inquirer.prompt(questions);
  if (answers) {
    let listId = null;
    for (let i = 0; i < lists.length; i++) {
      if (lists[i].title === answers.list) {
        answers.list = lists[i].id;
        break;
      }
    }
  }

  return answers;
};

const performFormAction = async (settings, removeCustomization) => {
  let jsomSettings = await (new JsomNode()).wizard();
  // console.log(settings);

  const configs = await getConfigs(settings);
  const configCtx = await setupPnp(configs);

  const ctx = SP.ClientContext.get_current();
  const web = ctx.get_web();
  const site = ctx.get_site();
  const rootWeb = site.get_rootWeb();
  const lists = web.get_lists();
  ctx.load(rootWeb, 'RootFolder');
  ctx.load(web, 'RootFolder');
  ctx.load(lists);
  await ctx.executeQueryPromise();

  let lds = [];
  for (let i = 0; i < lists.get_count(); i++) {
    let listData = lists.getItemAtIndex(i);
    if (!listData.get_hidden() && listData.get_baseType() !== SP.BaseType.documentLibrary && !listData.get_hidden()) {
      lds.push({
        title: listData.get_title(),
        id: listData.get_id().toString()
      });
    }

  }

  let userResults = await promptUser(lds);
  //console.log(userResults);

  const list = web.get_lists().getById(userResults.list);

  ctx.load(list,
    'DefaultEditFormUrl',
    'DefaultNewFormUrl',
    'DefaultDisplayFormUrl',
    'RootFolder',
    'DefaultViewUrl',
    'Title');
  await ctx.executeQueryPromise();

  let listFormsBase = list.get_defaultViewUrl().substring(0, list.get_defaultViewUrl().lastIndexOf('/'));
  //let possibleSurplusRootWebFragment = rootWeb.get_rootFolder().get_serverRelativeUrl();

  let forms =  [
    list.get_defaultEditFormUrl() == null ? listFormsBase + '/EditForm.aspx' : list.get_defaultEditFormUrl(),
    list.get_defaultNewFormUrl() == null ? listFormsBase + '/NewForm.aspx' : list.get_defaultNewFormUrl(),
    list.get_defaultDisplayFormUrl()  == null ? listFormsBase + '/DispForm.aspx' : list.get_defaultDisplayFormUrl()
  ];

  if (userResults.forms === 'New') {
    forms = [forms[1]];
  } else if (userResults.forms === 'Edit') {
    forms = [forms[0]];
  } else if (userResults.forms === 'Display') {
    forms = [forms[2]];
  }

  let webPartColleciton = [];
  try {
    for (let i = 0; i < forms.length; i++) {
      console.log(`Processing list '${list.get_title()}', form: ${forms[i]}`);
      let file = web.getFileByServerRelativeUrl(forms[i]);
      let webPartMngr = file.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
      let webparts = webPartMngr.get_webParts();
      ctx.load(webparts);
      webPartColleciton.push(webparts);
    }
    await ctx.executeQueryPromise();
  } catch (e) {
    console.error(e);
  }

  try {
    for (let i = webPartColleciton.length - 1; i >= 0; i--) {
      let wpDefCount = webPartColleciton[i].get_count();
      for (let j = wpDefCount - 1; j >= 0; j--) {
        let wp = webPartColleciton[i].get_item(j);
        let props = wp.get_webPart().get_properties();
        ctx.load(props);
        await ctx.executeQueryPromise();
        let objValues = props.get_fieldValues();
        // console.log(JSON.stringify(objValues));
        // console.log(objValues.Title);
        if (objValues && objValues.Title.match(/react form/gi)) {
          wp.deleteWebPart();
          await ctx.executeQueryPromise();
        } else {
          let webPart = wp.get_webPart();
          ctx.load(webPart);
          webPart.set_hidden(!removeCustomization);
          wp.saveWebPartChanges();
          await ctx.executeQueryPromise();
        }
      }
    }

  } catch(e) {
    console.error(e);
  }

  if (!removeCustomization) {
    let contentLink = configs.privateConf.siteUrl + '/' + configs.appConfig.spFolder + '/webparts/listform.cewp.html';
    var webPartXml =
      '<?xml version="1.0" encoding="utf-8"?>' +
        '<WebPart xmlns="http://schemas.microsoft.com/WebPart/v2">' +
            '<Assembly>Microsoft.SharePoint, Version=' + configCtx.envCode + '.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>' +
            '<TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>' +
            '<Title>React Form</Title>' +
            '<ContentLink xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor">' + contentLink + '</ContentLink>'+
            '<Description>$Resources:core,ContentEditorWebPartDescription;</Description>' +
            '<PartImageLarge>/_layouts/images/mscontl.gif</PartImageLarge>' +
        '</WebPart>';

    for (let i = 0; i < forms.length; i++) {
      let file = web.getFileByServerRelativeUrl(forms[i]);
      let webPartMngr = file.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
      let webPartDef = webPartMngr.importWebPart(webPartXml);
      let webPart = webPartDef.get_webPart();
      webPartMngr.addWebPart(webPart, 'Main', 1);

      ctx.load(webPart);
      await ctx.executeQueryPromise();
    }
  }

  // console.log(2)
  try {
    // list.set_defaultEditFormUrl(forms[0]);
    // list.set_defaultNewFormUrl(forms[1]);
    // list.set_defaultDisplayFormUrl(forms[2]);
    // await ctx.executeQueryPromise();
  } catch (e) {
    console.error(e);
  }
};

module.exports = customTask((gulp, $, settings) => {
  gulp.task('replaceListForm', cb => {
    (async () => {
     await performFormAction(settings, false);
    })().then(_ => cb()).catch(cb);
  });

  gulp.task('restoreListForm', cb => {
    (async () => {
     await performFormAction(settings, true);
    })().then(_ => cb()).catch(cb);
  });
});
