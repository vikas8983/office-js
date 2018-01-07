[![NPM Deployment Status](https://travis-ci.org/OfficeDev/office-js.svg?branch=release)](https://travis-ci.org/OfficeDev/office-js/builds)

# This branch acts as the "command" infrastructure for NPM deployments of "@microsoft/office-js"

It's used to trigger a deployment to one of the official branches.


# To trigger a deployment

1. Open https://github.com/OfficeDev/office-js/edit/deployment-queue/DEPLOY_REQUEST.yaml
2. Modify the parameters.
3. Submit your changes as a pull request.
4. Once the pull request is approved, the deployment will take place.


# More info on this branch

The branch is an orphan branch, NOT INTENDED TO BE MERGED OR PULLED-INTO from the normal branches.  It is here for the the triggering of deployments only.

The only folders/files in it are:

1. ".git" (because all folders have it)
2. "deployment.script" folder, where the entirety of the deployment script stuff resides
3. ".gitignore" THE ONLY IDENTICAL FILE to what is in the official branches.
4. ".travis.yml" for TravisCI deployment.
5. "DEPLOY_REQUEST.yaml", which specifies the deployment parameters (see above).
6. "README.md", this file.  Note that it is different from the official README.
