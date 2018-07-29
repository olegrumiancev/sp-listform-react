# sp-listform-react


### Bring Office UI Fabric-based forms into SharePoint 2013 / 2016 / Online
##### This project was intended as NodeJS learning ground, but quickly became an attempt to serve the awesome SharePoint developer community by providing a super easy solution that developers can use to transform OOTB ugly SharePoint forms into beautiful React-based forms.


##### Insipred by [Denis Molodtsov's](https://github.com/zerg00s/) [AngularForms](https://github.com/Zerg00s/AngularForms) and [spforms](https://github.com/Zerg00s/spforms) solutions that do lightweight form transformations in AngularJS.
##### The scaffolding is based on another awesome Microsoft MVP (which I think should be granted MVP for life), [Andrew Koltyakov's](https://github.com/koltyakov) yeoman generator, called [SharePoint Push'n'Pull](https://github.com/koltyakov/generator-sppp)



##### The working React components, such as **ListForm and FormField** and their interfaces are provided by a separate NPM package that I poured my soul into (so, if it's bad, don't be very harsh!), called **[sp-react-formfields](https://npmjs.com/package/sp-react-formfields)**
## Description
The output of this scaffolded solution will be (by default) in the ./dist folder and will contain webpacked *.JS / *.CSS and *.HTML files. The intended purpose of them is to use *.HTML files as content sources for Contenr Editor WebParts (CEWPs).

The structure of scaffolded solution is

![img](https://olegrumiancev.github.io/sp-listform-react/structure.JPG)

You are interested in 4 folders:
- **config** - this will contain information about your target site (collection), deployment folder in that site (like _catalogs/mycode/... or even just SiteAssets/mycode/... -- your choice!). **Fill config by calling NPM RUN CONFIG**
- **scripts** - open root.tsx file if you want to write custom UI for your form (and not just use default one line for one field kind of deal). Look for function called **renderCustomFieldLogic** and modify it to return your custom JSX. to create form fields user **FormField** component and pass internal name as string.
- **webparts** - this folder contains two Handlebars template files which if you look at them are just HTML inside with some parameters substituted during build process and it is out of them that we are getting resulting *.HTML files in the ./dist folder. Once ./dist is uploaded - (we have a **gulp task** for this!!) they will be ready to be referenced in the CEWPs
- **tools / tasks**  this contains two gulp tasks that I specifically wrote for "enhancing" and "de-enhancing" SharePoint list forms
   - call **gulp replaceListForm** to connect to your configured site, select a list and a Form Mode (New / Edit / Display / All) to transform
   - call **gulp restoreListForm** to connect to your configured site, select a list and Form Mode to revert back to SharePoint OOTB form

## Features
  - Readily available scaffolding solution
  - Fork, clone and with a few actions you will transform any SharePoint list which uses standard OOTB fields
  - Almost of the OOTB fields are supported:
    - Text
    - Note (Rich mode as well!)
    - Boolean
    - Attachments!
    - Choice, Multichoice
    - Lookup, Multilookup
    - User, Usermulti
    - URL / Picture
    - Number
    - Currency
    - DateTime
    - Taxonomy!
 - Fields are validation-aware:
   - They react to required setting
   - Number and curreny fields know about Min/Max settings
   - Text field reacts to 255 default char limit and to custom limit setting
  - User field correctly suggests principals only from a specific SharePoint group, if that setting is enabled
  - Lookup field renders values from correct lookup list field, if anything other than "Title" is in the settings
  - Attachment field is based on DropzoneJS and is basically a droppable area. But users can also click on the area to evoke a standard browser file select dialog

## Minimal path to awesome

##### 1. Fork / clone the solution
##### 2. Open code editor
##### 3. Run
...
```sh
$ npm install
$ npm run config
```

To build the solution use

```sh
$ npm run build (same as gulp build --prod, this will build in production mode)
or
$ gulp build (to build in dev mode)
```

To push your build files to a place in your target site
```sh
$ gulp push (will push all files)
or
$ gulp push --diff (to only push files that differ from what is in SP)
```

Useful command while developing - this will launch a background task that on every save of a file will run
build and push --diff tasks
```sh
$ gulp watch
```

### Transfomation example
![img](https://olegrumiancev.github.io/sp-listform-react/transform.gif)
