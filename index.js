#!/usr/bin/env node
/**
 * DOCX to Markdown Converter
 * 
 * Converts .docx and .doc files to Markdown format.
 * Uses pandoc if available, otherwise falls back to native XML parsing.
 * 
 * This is a local file processing tool - no API tokens required.
 */

const fs = require('fs');
const path = require('path');

// Parse command line arguments  
const args = process.argv.slice(2);

function parseArgs() {
  const parsed = {
    command: null,
    positional: [],
    options: {}
  };
  
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg.startsWith('-')) {
      if (arg.startsWith('--')) {
        const [key, value] = arg.slice(2).split('=', 2);
        if (value !== undefined) {
          parsed.options[key] = value;
        } else if (i + 1 < args.length && !args[i + 1].startsWith('-')) {
          parsed.options[key] = args[i + 1];
          i++;
        } else {
          parsed.options[key] = true;
        }
      } else {
        const flag = arg.slice(1);
        if (i + 1 < args.length && !args[i + 1].startsWith('-')) {
          parsed.options[flag] = args[i + 1];
          i++;
        } else {
          parsed.options[flag] = true;
        }
      }
    } else {
      if (parsed.command === null) {
        parsed.command = arg;
      } else {
        parsed.positional.push(arg);
      }
    }
  }
  
  return parsed;
}

/**
 * Execute a shell command and return stdout
 * Uses the sandbox's built-in exec if available, otherwise __ipc__
 */
function execCommand(cmd, options = {}) {
  // In sandbox environment, use __ipc__ for system commands
  if (typeof __ipc__ === 'function') {
    const result = __ipc__('system.exec', cmd);
    if (result.err) {
      throw new Error(result.err);
    }
    // Check for stderr errors if stdout is empty (command likely failed)
    if (!result.stdout && result.stderr && !options.ignoreStderr) {
      throw new Error(result.stderr);
    }
    return result.stdout || '';
  }
  
  // Fallback for non-sandbox environment
  throw new Error('System command execution not available in this environment');
}

/**
 * Check if pandoc is available
 */
function hasPandoc() {
  try {
    const result = execCommand('which pandoc');
    return result.trim().length > 0;
  } catch (e) {
    return false;
  }
}

/**
 * Resolve path properly (sandbox-compatible)
 */
function resolvePath(inputPath) {
  if (path.isAbsolute(inputPath)) {
    return inputPath;
  }
  // Use process.cwd() for relative paths (sandbox path.resolve is broken)
  return path.join(process.cwd(), inputPath);
}

/**
 * Convert DOCX to Markdown using pandoc
 */
function convertWithPandoc(inputPath, options = {}) {
  const absPath = resolvePath(inputPath);
  
  if (!fs.existsSync(absPath)) {
    throw new Error(`File not found: ${inputPath} (resolved to: ${absPath})`);
  }
  
  // Build pandoc command
  let cmd = `pandoc "${absPath}" -f docx -t markdown`;
  
  // Add options
  if (options.wrap === 'none' || options.nowrap) {
    cmd += ' --wrap=none';
  } else if (options.wrap) {
    cmd += ` --wrap=${options.wrap}`;
  }
  
  if (options.standalone) {
    cmd += ' -s';
  }
  
  if (options.extractMedia) {
    cmd += ` --extract-media="${options.extractMedia}"`;
  }
  
  if (options.reference) {
    cmd += ' --reference-links';
  }
  
  // ATX-style headers (# style instead of underline)
  if (options.atx !== false) {
    cmd += ' --markdown-headings=atx';
  }
  
  const result = execCommand(cmd);
  
  // Validate result - if empty, something went wrong
  if (!result || result.trim().length === 0) {
    throw new Error(`Conversion failed: pandoc returned no output for ${inputPath}. The file may be corrupted.`);
  }
  
  return result;
}

/**
 * Parse XML content to extract text and basic structure
 * This is a fallback when pandoc is not available
 */
function parseDocxXml(xmlContent) {
  const lines = [];
  let currentParagraph = '';
  let inParagraph = false;
  let isBold = false;
  let isItalic = false;
  let isHeading = false;
  let headingLevel = 0;
  
  // Very basic XML parsing for DOCX document.xml
  // This handles common patterns but won't cover all edge cases
  
  // Extract text from <w:t> tags
  const textMatches = xmlContent.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g);
  const paragraphMatches = xmlContent.split(/<w:p[^>]*>/);
  
  for (const para of paragraphMatches) {
    if (!para.trim()) continue;
    
    // Check for heading style
    const styleMatch = para.match(/<w:pStyle w:val="([^"]+)"/);
    let prefix = '';
    
    if (styleMatch) {
      const style = styleMatch[1].toLowerCase();
      if (style.includes('heading1') || style === 'h1') {
        prefix = '# ';
      } else if (style.includes('heading2') || style === 'h2') {
        prefix = '## ';
      } else if (style.includes('heading3') || style === 'h3') {
        prefix = '### ';
      } else if (style.includes('heading4') || style === 'h4') {
        prefix = '#### ';
      } else if (style.includes('title')) {
        prefix = '# ';
      }
    }
    
    // Extract text content
    let text = '';
    const textParts = para.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g);
    
    for (const match of textParts) {
      text += match[1];
    }
    
    if (text.trim()) {
      // Check for bold
      if (para.includes('<w:b/>') || para.includes('<w:b ')) {
        text = `**${text.trim()}**`;
      }
      
      // Check for italic
      if (para.includes('<w:i/>') || para.includes('<w:i ')) {
        text = `*${text.trim()}*`;
      }
      
      lines.push(prefix + text.trim());
    }
  }
  
  return lines.join('\n\n');
}

/**
 * Convert DOCX using native XML parsing (fallback method)
 */
function convertNative(inputPath) {
  const absPath = resolvePath(inputPath);
  
  if (!fs.existsSync(absPath)) {
    throw new Error(`File not found: ${inputPath}`);
  }
  
  // DOCX files are ZIP archives
  // We need to extract and parse word/document.xml
  
  // Use unzip to extract document.xml
  const tmpDir = `/tmp/docx2md_${Date.now()}`;
  
  try {
    execCommand(`mkdir -p "${tmpDir}"`);
    execCommand(`unzip -q "${absPath}" -d "${tmpDir}"`);
    
    const documentPath = `${tmpDir}/word/document.xml`;
    
    if (!fs.existsSync(documentPath)) {
      throw new Error('Invalid DOCX file: missing word/document.xml');
    }
    
    const xmlContent = fs.readFileSync(documentPath, 'utf8');
    const markdown = parseDocxXml(xmlContent);
    
    // Cleanup
    execCommand(`rm -rf "${tmpDir}"`);
    
    return markdown;
  } catch (error) {
    // Cleanup on error
    try {
      execCommand(`rm -rf "${tmpDir}"`);
    } catch (e) {}
    throw error;
  }
}

/**
 * Convert a DOCX file to Markdown
 */
function convert(inputPath, options = {}) {
  // Check file extension
  const ext = path.extname(inputPath).toLowerCase();
  
  if (ext !== '.docx' && ext !== '.doc') {
    throw new Error(`Unsupported file type: ${ext}. Only .docx and .doc files are supported.`);
  }
  
  if (ext === '.doc') {
    // .doc files require pandoc or LibreOffice
    if (!hasPandoc()) {
      throw new Error('.doc files require pandoc to be installed. Install with: brew install pandoc');
    }
  }
  
  // Try pandoc first (better quality)
  if (!options.native && hasPandoc()) {
    return convertWithPandoc(inputPath, options);
  }
  
  // Fallback to native parsing
  if (ext === '.docx') {
    return convertNative(inputPath);
  }
  
  throw new Error('Conversion failed: pandoc not available and native parsing only supports .docx');
}

/**
 * Get file info
 */
function getInfo(inputPath) {
  const absPath = resolvePath(inputPath);
  
  if (!fs.existsSync(absPath)) {
    throw new Error(`File not found: ${inputPath}`);
  }
  
  const stats = fs.statSync(absPath);
  const ext = path.extname(inputPath).toLowerCase();
  
  const info = {
    path: absPath,
    name: path.basename(inputPath),
    extension: ext,
    size: stats.size,
    sizeHuman: formatBytes(stats.size),
    modified: stats.mtime.toISOString(),
    pandocAvailable: hasPandoc()
  };
  
  // For DOCX, try to get document properties
  if (ext === '.docx') {
    try {
      const tmpDir = `/tmp/docx2md_info_${Date.now()}`;
      execCommand(`mkdir -p "${tmpDir}"`);
      execCommand(`unzip -q "${absPath}" -d "${tmpDir}"`);
      
      // Try to read core properties
      const propsPath = `${tmpDir}/docProps/core.xml`;
      if (fs.existsSync(propsPath)) {
        const propsXml = fs.readFileSync(propsPath, 'utf8');
        
        // Extract title
        const titleMatch = propsXml.match(/<dc:title>([^<]+)<\/dc:title>/);
        if (titleMatch) info.title = titleMatch[1];
        
        // Extract creator
        const creatorMatch = propsXml.match(/<dc:creator>([^<]+)<\/dc:creator>/);
        if (creatorMatch) info.creator = creatorMatch[1];
        
        // Extract created date
        const createdMatch = propsXml.match(/<dcterms:created[^>]*>([^<]+)<\/dcterms:created>/);
        if (createdMatch) info.created = createdMatch[1];
        
        // Extract word count from app properties
        const appPropsPath = `${tmpDir}/docProps/app.xml`;
        if (fs.existsSync(appPropsPath)) {
          const appXml = fs.readFileSync(appPropsPath, 'utf8');
          
          const wordsMatch = appXml.match(/<Words>(\d+)<\/Words>/);
          if (wordsMatch) info.wordCount = parseInt(wordsMatch[1]);
          
          const pagesMatch = appXml.match(/<Pages>(\d+)<\/Pages>/);
          if (pagesMatch) info.pageCount = parseInt(pagesMatch[1]);
        }
      }
      
      execCommand(`rm -rf "${tmpDir}"`);
    } catch (e) {
      // Ignore errors reading properties
    }
  }
  
  return info;
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function printHelp() {
  console.log(`
DOCX to Markdown Converter

USAGE:
  docx2md <command> [options]

COMMANDS:
  convert <file>              Convert DOCX/DOC to Markdown
  info <file>                 Show document information
  help                        Show this help message

CONVERT OPTIONS:
  -o, --output <file>         Output file (default: stdout)
  --wrap <mode>               Line wrapping: auto, none, preserve (default: auto)
  --nowrap                    Disable line wrapping (same as --wrap=none)
  --standalone                Include YAML front matter
  --extract-media <dir>       Extract images to directory
  --reference                 Use reference-style links
  --native                    Force native parsing (skip pandoc)

OUTPUT OPTIONS:
  --json                      Output as JSON (for info command)
  --summary                   Human-readable output

EXAMPLES:
  # Convert and print to stdout
  docx2md convert document.docx

  # Convert and save to file
  docx2md convert document.docx -o output.md

  # Convert without line wrapping
  docx2md convert document.docx --nowrap -o output.md

  # Extract images while converting
  docx2md convert document.docx --extract-media=./images -o output.md

  # Get document info
  docx2md info document.docx --summary

NOTES:
  - Uses pandoc for high-quality conversion if available
  - Falls back to native XML parsing for .docx if pandoc not found
  - .doc files require pandoc to be installed
  - Install pandoc: brew install pandoc (macOS) or apt install pandoc (Linux)
`);
}

function main() {
  const parsed = parseArgs();
  
  if (!parsed.command || parsed.command === 'help' || parsed.options.help || parsed.options.h) {
    printHelp();
    return;
  }
  
  try {
    switch (parsed.command) {
      case 'convert': {
        const inputFile = parsed.positional[0];
        if (!inputFile) {
          console.error('Error: Input file required');
          console.error('Usage: docx2md convert <file> [-o output.md]');
          process.exit(1);
        }
        
        const options = {
          wrap: parsed.options.wrap,
          nowrap: parsed.options.nowrap,
          standalone: parsed.options.standalone,
          extractMedia: parsed.options['extract-media'],
          reference: parsed.options.reference,
          native: parsed.options.native
        };
        
        const markdown = convert(inputFile, options);
        
        // Output
        const outputFile = parsed.options.o || parsed.options.output;
        if (outputFile) {
          // Ensure output directory exists (use mkdir -p for sandbox compatibility)
          const outputDir = path.dirname(outputFile);
          if (outputDir && outputDir !== '.') {
            try {
              execCommand(`mkdir -p "${outputDir}"`);
            } catch (e) {
              // Directory may already exist, ignore error
            }
          }
          
          fs.writeFileSync(outputFile, markdown);
          
          if (parsed.options.summary) {
            console.log(`Converted: ${inputFile} -> ${outputFile}`);
            console.log(`Output size: ${formatBytes(markdown.length)}`);
          }
        } else {
          // Print to stdout
          console.log(markdown);
        }
        break;
      }
      
      case 'info': {
        const inputFile = parsed.positional[0];
        if (!inputFile) {
          console.error('Error: Input file required');
          console.error('Usage: docx2md info <file>');
          process.exit(1);
        }
        
        const info = getInfo(inputFile);
        
        if (parsed.options.json) {
          console.log(JSON.stringify(info, null, 2));
        } else {
          console.log(`File: ${info.name}`);
          console.log(`Path: ${info.path}`);
          console.log(`Type: ${info.extension.toUpperCase().slice(1)} document`);
          console.log(`Size: ${info.sizeHuman}`);
          console.log(`Modified: ${info.modified}`);
          
          if (info.title) console.log(`Title: ${info.title}`);
          if (info.creator) console.log(`Author: ${info.creator}`);
          if (info.created) console.log(`Created: ${info.created}`);
          if (info.wordCount) console.log(`Words: ${info.wordCount}`);
          if (info.pageCount) console.log(`Pages: ${info.pageCount}`);
          
          console.log(`\nPandoc: ${info.pandocAvailable ? 'Available' : 'Not installed'}`);
        }
        break;
      }
      
      default:
        // If command looks like a file, assume 'convert' command
        if (parsed.command.endsWith('.docx') || parsed.command.endsWith('.doc')) {
          // Treat command as input file
          const inputFile = parsed.command;
          const options = {
            wrap: parsed.options.wrap,
            nowrap: parsed.options.nowrap,
            standalone: parsed.options.standalone,
            extractMedia: parsed.options['extract-media'],
            reference: parsed.options.reference,
            native: parsed.options.native
          };
          
          const markdown = convert(inputFile, options);
          
          const outputFile = parsed.options.o || parsed.options.output;
          if (outputFile) {
            // Ensure output directory exists (use mkdir -p for sandbox compatibility)
            const outputDir = path.dirname(outputFile);
            if (outputDir && outputDir !== '.') {
              try {
                execCommand(`mkdir -p "${outputDir}"`);
              } catch (e) {
                // Directory may already exist, ignore error
              }
            }
            fs.writeFileSync(outputFile, markdown);
            
            if (parsed.options.summary) {
              console.log(`Converted: ${inputFile} -> ${outputFile}`);
              console.log(`Output size: ${formatBytes(markdown.length)}`);
            }
          } else {
            console.log(markdown);
          }
        } else {
          console.error(`Error: Unknown command '${parsed.command}'`);
          console.error('\nRun: docx2md help');
          process.exit(1);
        }
    }
    
  } catch (error) {
    console.error(`Error: ${error.message}`);
    if (process.env.DEBUG) {
      console.error('Stack trace:', error.stack);
    }
    process.exit(1);
  }
}

// Execute
main();
