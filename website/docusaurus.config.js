module.exports={
  "title": "PptxGenJS",
  "tagline": "Create JavaScript PowerPoint Presentations",
  "url": "https://gitbrent.github.io",
  "baseUrl": "/PptxGenJS/",
  "organizationName": "PptxGenJS",
  "projectName": "PptxGenJS",
  "scripts": [
    "https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/dist/pptxgen.bundle.js",
    "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/9.12.0/highlight.min.js"
  ],
  "favicon": "img/favicon.png",
  "customFields": {
    "repoUrl": "https://github.com/gitbrent/PptxGenJS"
  },
  "onBrokenLinks": "log",
  "onBrokenMarkdownLinks": "log",
  "presets": [
    [
      "@docusaurus/preset-classic",
      {
        "docs": {
          "showLastUpdateAuthor": true,
          "showLastUpdateTime": true,
          "path": "./docs",
          "sidebarPath": "./sidebars.json"
        },
        "blog": {},
        "theme": {
          "customCss": "../src/css/customTheme.css"
        }
      }
    ]
  ],
  "plugins": [
    [
      "@docusaurus/plugin-client-redirects",
      {
        "fromExtensions": [
          "html"
        ]
      }
    ]
  ],
  "themeConfig": {
    "navbar": {
      "title": "PptxGenJS",
      "logo": {
        "src": "img/pptxgenjs.svg"
      },
      "items": [
        {
          "href": "https://github.com/gitbrent/PptxGenJS/releases",
          "label": "Download",
          "position": "left"
        },
        {
          "to": "docs/",
          "label": "Get Started",
          "position": "left"
        },
        {
          "to": "docs/installation",
          "label": "API Documentation",
          "position": "left"
        },
        {
          "href": "https://github.com/gitbrent/PptxGenJS/",
          "label": "GitHub",
          "position": "left"
        }
      ]
    },
    "footer": {
      "links": [],
      "copyright": "Copyright © 2021 Brent Ely",
      "logo": {
        "src": "img/pptxgenjs.svg"
      }
    },
    "gtag": {
      "trackingID": "UA-75147115-1"
    }
  }
}
