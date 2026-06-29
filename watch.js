const chokidar = require('chokidar');
const { spawn } = require('child_process');
const path = require('path');

console.log('🚀 Manufacturing Analytics - File Watcher Started');
console.log('📁 Watching Excel files in ./data folder');
console.log('⏱️ Auto-rebuild on changes...\n');

// Watch for Excel file changes
const watcher = chokidar.watch('./data/*.xlsx', {
  persistent: true,
  awaitWriteFinish: {
    stabilityThreshold: 2000,
    pollInterval: 100
  }
});

function rebuildAnalytics() {
  console.log('\n📊 Excel files changed - rebuilding analytics...');
  
  const python = spawn('python3', ['process_analytics.py']);
  
  python.stdout.on('data', (data) => {
    console.log(data.toString());
  });
  
  python.stderr.on('data', (data) => {
    console.error(`Error: ${data}`);
  });
  
  python.on('close', (code) => {
    if (code === 0) {
      console.log('✅ Analytics data updated successfully');
      console.log(`⏰ Ready for next changes at ${new Date().toLocaleTimeString()}\n`);
    } else {
      console.error('❌ Build failed');
    }
  });
}

// Events
watcher.on('add', (file) => {
  console.log(`✨ New file detected: ${path.basename(file)}`);
  rebuildAnalytics();
});

watcher.on('change', (file) => {
  console.log(`📝 Modified: ${path.basename(file)}`);
  rebuildAnalytics();
});

watcher.on('unlink', (file) => {
  console.log(`🗑️  Deleted: ${path.basename(file)}`);
});

watcher.on('error', (error) => {
  console.error(`Watcher error: ${error}`);
});

// Initial build
console.log('🔨 Running initial build...\n');
rebuildAnalytics();

// Keep process alive
process.on('SIGINT', () => {
  console.log('\n👋 Stopping file watcher...');
  watcher.close();
  process.exit(0);
});
