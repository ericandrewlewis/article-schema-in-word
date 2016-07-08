# Article Schema in Microsoft Word

## Local dev

Install node.js

[Install relevant node things in this tutorial](http://dev.office.com/docs/add-ins/get-started/create-an-office-add-in-using-any-editor)

Install dependencies

```
npm install && bower install
```

Serve the files locally

```
gulp serve-static
```

Navigate to `https://localhost:8443/`, get the cert and add it to your computer/browser's keychain.

Go to a Word Online document.

In the menu bar click Insert > Add-Ins. Upload `manifest.xml`.

## See also

[The Word JavaScript API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)