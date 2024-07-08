# Performance Testing Using Lighthouse
Automated testing of lighthouse typically aims to ensure that web applications or websites meet certain performance, accessibility, best practices, and SEO standards.

## 🔰 Getting Started
#### - Prerequisite: Node.js should be installed

#### Steps:

1. Clone or Download the repo
2. Open Project Terminal on the main directory
4. Run `npm i` or `npm install` - To install all the dependencies(Must only be done once) 
5. Run `npm i -g lighthouse` or `npm install -g lighthouse` - To install lighthouse globally
6. Run `npm run lighthouse:dev` or `npm run lighthouse:sit` or `npm run lighthouse:uat` or `npm run lighthouse:prod` to start running the project

## ▶️ Run a tests:
In **package.json**, You can find the collection of **scripts** member -> an object hash that contains commands to be executed at various points in the lifecycle of your package. These scripts are typically run using `npm` or `yarn` commands.
#### Example:
* **"npm run lighthouse:dev"**: "node features/script/performanceLighthouseScore.js DEV"
    * This script runs when you execute lighthouse in DEV
* **"npm run lighthouse:sit"**: "node features/script/performanceLighthouseScore.js SIT"
    * This script runs when you execute lighthouse in SIT
* **"npm run lighthouse:uat"**: "node features/script/performanceLighthouseScore.js UAT"
    * This script runs when you execute lighthouse in UAT
* **"npm run lighthouse:prod"**: "node features/script/performanceLighthouseScore.js PROD"
    * This script runs when you execute lighthouse in PROD

## ⚙️ Build With:
- [Selenium](https://github.com/SeleniumHQ/selenium) - Automation Framework.
