#!/usr/bin/env node

/**
 * Smart dev mode for js-to-ppt
 * 
 * Usage: npm run dev
 * 
 * What it does:
 * 1. Auto-finds projects that use @linedotai/js-to-ppt in sibling folders
 * 2. Links all found projects to use this local package
 * 3. Saves original lock files (package-lock.json, yarn.lock)
 * 4. Runs Rollup watch mode
 * 5. When you stop (Ctrl+C):
 *    - Restores all projects to their published version
 *    - Restores original lock files
 */

const { spawn, execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

const PACKAGE_NAME = '@linedotai/js-to-ppt';

// Find projects by checking sibling directories
function findProjects() {
  const packageDir = __dirname;
  const parentDir = path.dirname(packageDir);
  
  // Projects to look for
  const projectConfigs = [
    {
      names: ['linedot-backend', 'flyingshelf-backend', 'backend', 'api'],
      label: 'Backend'
    },
    {
      names: ['linedot-studio', 'flyingshelf-studio', 'flyingshelf', 'studio', 'frontend', 'app', 'linedot-app'],
      label: 'Studio'
    },
    {
      names: ['linedot-photographer', 'flyingshelf-photographer', 'photographer'],
      label: 'Photographer'
    },
    {
      names: ['convert-to-ppt'],
      label: 'Convert-to-PPT'
    }
  ];
  
  const foundProjects = [];
  
  for (const config of projectConfigs) {
    for (const name of config.names) {
      const candidatePath = path.join(parentDir, name);
      const packageJsonPath = path.join(candidatePath, 'package.json');
      
      if (fs.existsSync(packageJsonPath)) {
        try {
          const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
          // Check if it has our package
          if (pkg.dependencies?.[PACKAGE_NAME] || pkg.devDependencies?.[PACKAGE_NAME]) {
            foundProjects.push({
              path: candidatePath,
              label: config.label,
              name: name
            });
            break; // Found this project, move to next config
          }
        } catch (e) {
          // Invalid package.json, skip
        }
      }
    }
  }
  
  return foundProjects;
}

console.log('\n🚀 Starting dev mode for js-to-ppt...\n');

// Build first
console.log('📦 Building package...\n');
try {
  execSync('npm run build', { stdio: 'inherit', cwd: __dirname });
  console.log('\n✅ Build complete\n');
} catch (error) {
  console.error('❌ Build failed\n');
  process.exit(1);
}

const projects = findProjects();

if (projects.length === 0) {
  console.log('ℹ️  No projects found that use ' + PACKAGE_NAME);
  console.log('   Running watch-only mode (no auto-linking)\n');
  
  // Just run watch mode
  const watch = spawn('npm', ['run', 'watch'], { stdio: 'inherit', shell: true, cwd: __dirname });
  process.exit(0);
}

console.log(`📁 Found ${projects.length} project(s) using ${PACKAGE_NAME}:`);
projects.forEach(p => console.log(`   - ${p.label} (${p.name})`));
console.log('');

// Store project states for cleanup
const projectStates = [];

// Link each project to local package
for (const project of projects) {
  const packageJsonPath = path.join(project.path, 'package.json');
  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
  const currentVersion = packageJson.dependencies?.[PACKAGE_NAME] || packageJson.devDependencies?.[PACKAGE_NAME];
  const wasAlreadyLocal = currentVersion && currentVersion.startsWith('file:');
  
  // Save lock files before making changes
  const lockFiles = {};
  const packageLockPath = path.join(project.path, 'package-lock.json');
  const yarnLockPath = path.join(project.path, 'yarn.lock');
  
  if (fs.existsSync(packageLockPath)) {
    lockFiles.packageLock = fs.readFileSync(packageLockPath, 'utf8');
  }
  if (fs.existsSync(yarnLockPath)) {
    lockFiles.yarnLock = fs.readFileSync(yarnLockPath, 'utf8');
  }
  
  projectStates.push({
    project,
    originalVersion: currentVersion,
    wasAlreadyLocal,
    lockFiles,
    packageJsonPath
  });
  
  // Link to local if not already
  if (!wasAlreadyLocal) {
    console.log(`🔗 Linking ${project.label} to LOCAL js-to-ppt...`);
    console.log(`   (will restore to "${currentVersion}" on exit)\n`);
    
    const relativePath = path.relative(project.path, __dirname);
    
    try {
      execSync(`cd "${project.path}" && npm install file:${relativePath}`, { stdio: 'inherit' });
      console.log(`✅ ${project.label} now using LOCAL js-to-ppt\n`);
    } catch (error) {
      console.error(`❌ Failed to link ${project.label}\n`);
    }
  } else {
    console.log(`✅ ${project.label} already using LOCAL js-to-ppt\n`);
  }
}

console.log('👀 Starting Rollup watch mode...');
console.log('💡 Press Ctrl+C to stop and restore all projects\n');

// Start watch mode
const watchProcess = spawn('npm', ['run', 'watch'], { stdio: 'inherit', shell: true, cwd: __dirname });

// Cleanup on exit
let isCleaningUp = false;

const cleanup = () => {
  if (isCleaningUp) return;
  isCleaningUp = true;
  
  console.log('\n\n🛑 Stopping dev mode...\n');
  
  // Restore each project
  for (const state of projectStates) {
    const { project, originalVersion, wasAlreadyLocal, lockFiles, packageJsonPath } = state;
    
    // Only unlink if we linked it (not if it was already local)
    if (!wasAlreadyLocal && fs.existsSync(packageJsonPath)) {
      console.log(`🔄 Restoring ${project.label} to original version: ${originalVersion}...`);
      
      try {
        // Read current package.json
        const currentPkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
        
        // Restore original version in the exact same location (dependencies or devDependencies)
        if (currentPkg.dependencies?.[PACKAGE_NAME]) {
          currentPkg.dependencies[PACKAGE_NAME] = originalVersion;
        } else if (currentPkg.devDependencies?.[PACKAGE_NAME]) {
          currentPkg.devDependencies[PACKAGE_NAME] = originalVersion;
        }
        
        // Write back with proper formatting
        fs.writeFileSync(packageJsonPath, JSON.stringify(currentPkg, null, 2) + '\n', 'utf8');
        
        // Restore lock files BEFORE running npm install
        const packageLockPath = path.join(project.path, 'package-lock.json');
        const yarnLockPath = path.join(project.path, 'yarn.lock');
        
        if (lockFiles.packageLock) {
          console.log(`   Restoring package-lock.json...`);
          fs.writeFileSync(packageLockPath, lockFiles.packageLock, 'utf8');
        }
        if (lockFiles.yarnLock) {
          console.log(`   Restoring yarn.lock...`);
          fs.writeFileSync(yarnLockPath, lockFiles.yarnLock, 'utf8');
        }
        
        // Run npm install to update node_modules
        execSync(`cd "${project.path}" && npm install`, { stdio: 'inherit' });
        
        console.log(`✅ ${project.label} restored to original state\n`);
      } catch (error) {
        console.error(`⚠️  Failed to restore ${project.label}. Run manually:`);
        console.error(`   Edit ${project.name}/package.json and set ${PACKAGE_NAME} to ${originalVersion}\n`);
      }
    }
  }
  
  console.log('👋 Dev mode stopped\n');
  process.exit(0);
};

process.on('SIGINT', cleanup);
process.on('SIGTERM', cleanup);

watchProcess.on('exit', cleanup);
