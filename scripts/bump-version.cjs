const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const packageJsonPath = path.join(rootDir, 'package.json');
const packageSolutionPath = path.join(rootDir, 'config', 'package-solution.json');

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function writeJson(filePath, value) {
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function bumpPatchVersion(version) {
  const versionMatch = /^(\d+)\.(\d+)\.(\d+)$/.exec(version);

  if (!versionMatch) {
    throw new Error(`Invalid package.json version "${version}". Expected format: major.minor.patch`);
  }

  const major = Number(versionMatch[1]);
  const minor = Number(versionMatch[2]);
  const patch = Number(versionMatch[3]) + 1;

  return `${major}.${minor}.${patch}`;
}

function toSolutionVersion(packageVersion) {
  const versionMatch = /^(\d+)\.(\d+)\.(\d+)$/.exec(packageVersion);

  if (!versionMatch) {
    throw new Error(`Invalid package version "${packageVersion}" for solution mapping.`);
  }

  return `${versionMatch[1]}.${versionMatch[2]}.${versionMatch[3]}.0`;
}

const packageJson = readJson(packageJsonPath);
const packageSolution = readJson(packageSolutionPath);

const nextPackageVersion = bumpPatchVersion(packageJson.version);
const nextSolutionVersion = toSolutionVersion(nextPackageVersion);

packageJson.version = nextPackageVersion;
packageSolution.solution.version = nextSolutionVersion;

writeJson(packageJsonPath, packageJson);
writeJson(packageSolutionPath, packageSolution);

console.log(`Bumped versions -> package.json: ${nextPackageVersion}, package-solution.json: ${nextSolutionVersion}`);
