# 👩‍💻 Contributing to the _SharePoint Notifications_

This project welcomes contributions and suggestions. Thank you for your interest in contributing to _SharePoint Notifications_. In this guide, we will walk you through the steps to get started.

## 👉 Before you start

In order to help us process your contributions, please make sure you do the following:

- don't surprise us with big PR's. Instead _create an issue_ & start a discussion so we can agree on a direction before you invest a large amount of time.
- create your branch from `dev` (NOT `main`). This will make it easier for us to merge your changes.
- submit PR to the `dev` branch of this repo (NOT `main`), by either contributing to:
    - `spfx` folder contains the SPFx extension
    - `backend` folder contains the backend services (Azure Functions)

  PRs submitted to other branches will be declined.
- let us know what's in the PR: sometimes code is not enough and in order to help us understand your awesome work please follow the PR template to provide required information.
- don't commit code you didn't write.

Do not be afraid to ask question. We are here to help you succeed in helping us making a better product.

## 👣 How to start - Minimal path to awesome

> [!IMPORTANT]
> Before you start, familiarize yourself with the SPFx extensions development process or the Azure Functions development process. Make sure that your local dependencies on you developer machine matches the appropriate requirements.

- Fork this project. When creating the fork unselect the checkbox 'Copy the `main` branch only' to get your copy of all the existing branches (`dev`, `dev-spfx`, `dev-backend` and `main` branch).
- Clone the forked repository
- Open the project in Visual Studio Code
- Navigate to the folder of your choice: `backend` or `spfx`
- run `npm install`
- Depending on your currend working device...
    - `backend`: Press `F5` to run the Azure Functions locally and start debugging
    - `spfx`: Press `heft start` to start the project on your local machine
- Deploy the (managed) solution under `power-platform` to a Power Platform environment of your choice

## ❓ More guidance and tips

For more usage guidance and tips go to the repo [README](./Readme.md).