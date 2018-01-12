#!/usr/bin/env node --harmony

import { isString, isNil } from "lodash";
import * as chalk from 'chalk';
import * as shell from 'shelljs';
import * as fs from "fs-extra";
import * as path from 'path';
import * as jsyaml from 'js-yaml';

import { banner, stripSpaces, execCommand } from "./util";
import * as VersionUtils from "./version-number-utils";

declare var process: {
    env: IEnvironmentVariables
    exit: (status: number) => void;
};

const TRAVIS_AUTO_COMMIT_TEXT = "[TRAVIS CI AUTO-COMMIT]";
const TOKENIZED_GITHUB_PUSH_URL = `https://<<<token>>>@github.com/OfficeDev/office-js.git`;
const DEPLOYMENT_YAML_FILENAME = "NPM.DEPLOYMENT.INFO.yaml";
const DEPLOY_REQUEST_FILENAME = "DEPLOY_REQUEST.yaml";

const REQUIRED_ADDITIONAL_FIELDS: Array<keyof IEnvironmentVariables> = ['GH_TOKEN'];

interface IEnvironmentVariables {
    TRAVIS: string,
    TRAVIS_BRANCH: string,
    TRAVIS_PULL_REQUEST: string,
    TRAVIS_COMMIT: string,
    TRAVIS_COMMIT_MESSAGE: string,
    TRAVIS_BUILD_ID: string,
    TRAVIS_BUILD_NUMBER: string,
    TRAVIS_BUILD_DIR: string,

    /**
     * GitHub token generated using https://github.com/settings/tokens,
     *     bearing permissions for "repo:status", "repo_deployment", and "public_repo".
     * This is a personal access token, so the commits always happen on behalf
     *     of the person who created the token.
     * The token is then entered as a hidden value in https://travis-ci.org/OfficeDev/office-js/settings */
    GH_TOKEN: string,

    /** A token for publishing to NPM.  It can be generated using "npm token create"
     * Note that you'll need NPM version 5.5.1+ to run this command.
     * https://docs.npmjs.com/getting-started/working_with_tokens
    */
    NPM_TOKEN: string
}

const OFFICIAL_BRANCHES = ["release", "release-next", "beta", "beta-next"];
const DEPLOYMENT_QUEUE_BRANCH = "deployment-queue";
const ADHOC_BRANCH_PREFIX = "__adhoc";
const DEFAULT_ADHOC_TAG = "adhoc";

interface IDeploymentInfoFromSubmittedRepo {
    tag: string,
    history: {
        commitMessage: string,
        adhocBranchName: string,
        fullCommitHistory: string
    }
}

const WORKING_DIRECTORY = path.resolve(process.env.TRAVIS_BUILD_DIR, "..", "working-travis-output-dir");

interface IOfficialBranchDeployRequest {
    targetBranch: string;
    from: string;
    deleteAdhocBranchOnSuccessfulDeployment: boolean;
}

interface IDeploymentParams {
    npmPublishTag: string;
    version: string;

    /** A script to run after cloning. Note that the current working directory at that point
     * is still the original one from the start of the script */
    afterCloneBeforeCommit?: (repoLocalFolderPath: string) => Promise<any>;

    deploymentInfoFromSubmittedRepoState: IDeploymentInfoFromSubmittedRepo;
}

(async () => {
    try {
        await attemptDeployScript();
        process.exit(0);
    }
    catch (error) {
        banner('AN ERROR OCCURRED', error.message || error, chalk.bold.red);
        console.error(error);

        banner('DEPLOYMENT DID NOT GET TRIGGERED', null, chalk.bold.red);
        process.exit(1);
    }
})();

async function attemptDeployScript() {
    printBuildStartInfo();
    makeWorkingDirectory();

    precheckOrExit();

    if (process.env.TRAVIS_BRANCH.startsWith(ADHOC_BRANCH_PREFIX)) {
        const deploymentInfoFromSubmittedRepoState = getDeploymentInfoFromSubmittedRepoState();
        const npmPublishTag = deploymentInfoFromSubmittedRepoState.tag;
        const version = await VersionUtils.getNextVersionNumberForNonReleaseTag(npmPublishTag);
        await doDeployment({ version, npmPublishTag, deploymentInfoFromSubmittedRepoState });
        return;

    } else if (process.env.TRAVIS_BRANCH === DEPLOYMENT_QUEUE_BRANCH) {
        await doOfficialDeployment();
        return;

    } else if (OFFICIAL_BRANCHES.indexOf(process.env.TRAVIS_BRANCH) >= 0) {
        const message = stripSpaces(`
                Deployment to one of the official branches must happen through the
                "${DEPLOYMENT_QUEUE_BRANCH}" branch. Please see
                https://github.com/OfficeDev/office-js/blob/deployment-queue/README.md
                for more info.
            `);
        banner('SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
        return;

    } else {
        const message = stripSpaces(`
            UNKNOWN BRANCH: Branch "${process.env.TRAVIS_BRANCH}" does not match any of the following:
                * A branch that starts with "__${ADHOC_BRANCH_PREFIX}".
                * The "${DEPLOYMENT_QUEUE_BRANCH}" branch.
                * Any of the following official branches: [${OFFICIAL_BRANCHES.map(item => `"${item}"`).join(", ")}].
        `);
        banner('SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
        return;
    }
}

function makeWorkingDirectory() {
    if (!fs.existsSync(WORKING_DIRECTORY)) {
        fs.mkdirSync(WORKING_DIRECTORY);
    } else {
        fs.emptyDirSync(WORKING_DIRECTORY);
    }

    banner("Working directory", WORKING_DIRECTORY);
}

function printBuildStartInfo() {
    const fieldsToPrint: (keyof IEnvironmentVariables)[] = [
        "TRAVIS",
        "TRAVIS_BRANCH",
        "TRAVIS_BUILD_ID",
        "TRAVIS_BUILD_NUMBER",
        "TRAVIS_COMMIT_MESSAGE",
        "TRAVIS_PULL_REQUEST",

        // "TRAVIS_BUILD_DIR":  Intentionally *NOT* outputting it,
        // since it serves no use to see, but causes issues if you copy-paste
        // the output of these Travis parameters from the log into "launch.json"
    ];

    const fieldsString = fieldsToPrint
        .map(item => `"${item}": "${process.env[item]}"`)
        .join(",\n");

    banner('TravisCI build started', fieldsString, chalk.green.bold);
}

function precheckOrExit(): void {
    /* Check if the code is running inside of travis.ci. If not abort immediately. */
    if (!process.env.TRAVIS) {
        banner('Deployment skipped', 'Not running inside of Travis.', chalk.yellow.bold);
        process.exit(0);
    }

    if (process.env.TRAVIS_COMMIT_MESSAGE && process.env.TRAVIS_COMMIT_MESSAGE.startsWith(TRAVIS_AUTO_COMMIT_TEXT)) {
        banner('Deployment skipped',
            `Skipping builds for commit messages labeled as "${TRAVIS_AUTO_COMMIT_TEXT}"`,
            chalk.yellow.bold);
        // NOTE: in practice, such builds should also be skipped because they'll have the text "skip ci" in them.
        // But this serves as a double-guarantee.
        process.exit(0);
    }

    // Careful! Need this check because otherwise, a pull request against master would immediately trigger a deployment.
    if (process.env.TRAVIS_PULL_REQUEST !== 'false') {
        banner('Deployment skipped', 'Skipping deploy for pull requests.', chalk.yellow.bold);
        process.exit(0);
    }

    REQUIRED_ADDITIONAL_FIELDS.forEach(key => {
        if (!isString(process.env[key]) || (process.env[key] as string).trim().length <= 0) {
            throw new Error(`"${key}" is a required global variables.`);
        }
    });
}

async function doDeployment(params: IDeploymentParams): Promise<void> {
    const { version, npmPublishTag, deploymentInfoFromSubmittedRepoState } = params;
    const gitTagName = "v" + params.version;

    banner("This deployment's target NPM version", "Target package version: " + version, chalk.magenta.bold);

    const deploymentFileContents = VersionUtils.generateDeploymentYamlText({
        npmPublishTag,
        version,
        historyInfo: deploymentInfoFromSubmittedRepoState.history,
        travisBuildId: process.env.TRAVIS_BUILD_ID,
        travisBuildNumber: process.env.TRAVIS_BUILD_NUMBER,
        includeTagUrls: npmPublishTag !== DEFAULT_ADHOC_TAG,
    });

    const repoLocalFolderPath = WORKING_DIRECTORY + "/" + "office-js/";
    fs.removeSync(repoLocalFolderPath);

    execCommand(`git clone ${TOKENIZED_GITHUB_PUSH_URL} ${repoLocalFolderPath}`, {
        token: process.env.GH_TOKEN
    });

    shell.pushd(repoLocalFolderPath);


    execCommand(`git checkout ${process.env.TRAVIS_BRANCH}`);
    execCommand('git config --add user.name "Travis CI"');
    execCommand('git config --add user.email "travis.ci@microsoft.com"');

    shell.popd();


    if (params.afterCloneBeforeCommit) {
        await params.afterCloneBeforeCommit(repoLocalFolderPath);
    }


    shell.pushd(repoLocalFolderPath);

    fs.writeFileSync(DEPLOYMENT_YAML_FILENAME, deploymentFileContents);
    execCommand(`git add ${DEPLOYMENT_YAML_FILENAME}`);

    VersionUtils.updatePackageJson(version);

    const commitMessage = `${TRAVIS_AUTO_COMMIT_TEXT} ${process.env.TRAVIS_COMMIT_MESSAGE} [skip ci]`;
    // Note: "skip CI" will skip travis running on the build.  https://docs.travis-ci.com/user/customizing-the-build/#Skipping-a-build

    execCommand(`git commit --allow-empty -m "${commitMessage}"`);
    execCommand(`git push`);


    // Now that the repo is updated, publish to NPM:

    fs.writeFileSync(".npmrc", `//registry.npmjs.org/:_authToken=${process.env.NPM_TOKEN}`);
    execCommand(`npm publish --tag ${npmPublishTag}`);

    // For us, the "release" tag is same as "latest" -- so for release, publish without a tag (implicit latest) too:
    if (npmPublishTag === "release") {
        execCommand(`npm publish`);
    }


    // If NPM succeeded, tag it and also add an NPM release:
    console.log(`Also tag the branch, and make a GitHub release: https://github.com/OfficeDev/office-js/releases/tag/${gitTagName}`);

    execCommand(`git tag -a ${gitTagName} -m "${commitMessage}"`);
    execCommand(`git push origin ${gitTagName}`);

    console.log(`FYI, if will need to delete the tag, run`);
    console.log(`    git push --delete origin {${gitTagName}}`);
    console.log(`And also discard the resulting draft from "https://github.com/OfficeDev/office-js/releases"`);

    const markdownReleasesNotes = VersionUtils.generateMarkdownDescription({
        npmPublishTag,
        DEPLOYMENT_YAML_FILENAME,
        version,
        commitMessage: deploymentInfoFromSubmittedRepoState.history.commitMessage,
        travisBuildId: process.env.TRAVIS_BUILD_ID,
        includeTagUrls: npmPublishTag !== DEFAULT_ADHOC_TAG,
    });

    // Documentation: https://developer.github.com/v3/repos/releases/#create-a-release
    const response = await fetch("https://api.github.com/repos/OfficeDev/office-js/releases", {
        method: "POST",
        headers: new Headers({
            "Authorization": `token ${process.env.GH_TOKEN}`
        }),
        body: JSON.stringify({
            "tag_name": gitTagName,
            "name": gitTagName,
            "body": markdownReleasesNotes,
            "prerelease": true,
            "draft": false
        })
    });

    if (response.status !== 201) {
        throw new Error(`Failed to create GitHub release; ${response.status}: ${response.statusText}`);
    }


    shell.popd();

    let removeLocalFolderAtCompletion = true;
    if (removeLocalFolderAtCompletion) {
        fs.removeSync(repoLocalFolderPath);
    }

    banner('SUCCESS, DEPLOYMENT COMPLETE!', markdownReleasesNotes.replace(/&nbsp;/g, ' '), chalk.green.bold);

    banner(`GitHub Releases page for v${version}`, `https://github.com/OfficeDev/office-js/releases/tag/v${version}`, chalk.green.bold);
}

function getDeploymentInfoFromSubmittedRepoState(): IDeploymentInfoFromSubmittedRepo {
    const fullPath = process.env.TRAVIS_BUILD_DIR + "/" + DEPLOYMENT_YAML_FILENAME;
    if (!fs.existsSync(fullPath)) {
        throw new Error(`No ${DEPLOYMENT_YAML_FILENAME} found!`);
    }

    const contents = fs.readFileSync(fullPath).toString();
    const result = jsyaml.safeLoad(contents) as IDeploymentInfoFromSubmittedRepo;
    if (result.history && result.tag) {
        return result;
    } else {
        throw new Error(`Missing required fields from in-repo "${DEPLOYMENT_YAML_FILENAME}" file.`);
    }
}

async function doOfficialDeployment(): Promise<void> {
    console.log(`First off: is there a request for a "targetBranch" and "from" in the ${DEPLOY_REQUEST_FILENAME} file?`);
    let currentYaml: IOfficialBranchDeployRequest = jsyaml.safeLoad(
        fs.readFileSync(process.env.TRAVIS_BUILD_DIR + "/" + DEPLOY_REQUEST_FILENAME).toString());

    if (isNil(currentYaml.targetBranch) || isNil(currentYaml.from)) {
        banner('SKIPPING DEPLOYMENT', `Nothing to deploy: missing "targetBranch" and/or "from" parameters.`, chalk.yellow.bold);
        return;
    }

    banner("DEPLOYMENT REQUEST DETECTED", stripSpaces(`
        Acknowledging request to deploy to
            "${currentYaml.targetBranch}"
        from
            ${currentYaml.from}
    `));

    banner('NOT READY TO DO OFFICIAL DEPLOYMENTS YET');
}

// async function getTravisQueueBranchDeploymentParams(): Promise<IDeploymentParams> {
//     const targetBranchName = process.env.TRAVIS_BRANCH.substr(TRAVIS_QUEUE_BRANCH_PREFIX.length);
//     if (OFFICIAL_BRANCHES.indexOf(targetBranchName) < 0) {
//         throw new Error(`The branch "${process.env.TRAVIS_BRANCH}" starts with the "${TRAVIS_QUEUE_BRANCH_PREFIX}" prefix, ` +
//             `but does not seem to match one of the official repo branches (${OFFICIAL_BRANCHES.map(item => `"${item}"`).join(", ")})`);
//     }

//     banner("Kicking off deployment to an official branch", `Target deployment branch: ${targetBranchName}`);

//     const version = (targetBranchName === "release")
//         ? await VersionUtils.getNextReleaseVersionNumber()
//         : await VersionUtils.getNextVersionNumberForNonReleaseTag(targetBranchName);

//     // Note that for official branches, the NPM tag and github tag are the same thing
//     //     (unlike for adhoc branches, where the tag is always "adhoc", but the branch is "__adhoc-xyz")
//     const npmPublishTag = targetBranchName;

//     return {
//         version,
//         npmPublishTag,

//         afterCloneBeforeCommit: async (repoLocalFolderPath: string) => {
//             console.log(`Delete all files except ".git" and the "dist" folder`);
//             fs.readdirSync(repoLocalFolderPath)
//                 .filter(filename => [".git", "dist"].indexOf(filename) < 0)
//                 .forEach(filename => fs.removeSync(repoLocalFolderPath + '/' + filename));

//             const repoCopyFolderPath = process.env.TRAVIS_BUILD_DIR + "/" + "office-js-repo-copy-release-branch/";
//             execCommand(`git clone ${TOKENIZED_GITHUB_PUSH_URL} ${repoCopyFolderPath}`, {
//                 token: process.env.GH_TOKEN
//             });

//             console.log(`Now, from a different clone of the "release" branch, copy all files except ".git" and the "dist" folder`);
//             fs.readdirSync(repoCopyFolderPath)
//                 .filter(filename => [".git", "dist"].indexOf(filename) < 0)
//                 .forEach(filename => fs.copySync(
//                     repoCopyFolderPath + '/' + filename,
//                     repoLocalFolderPath + '/' + filename,
//                     {
//                         preserveTimestamps: true
//                     }
//                 ));

//             fs.removeSync(repoCopyFolderPath);
//         }

//     };
// }
