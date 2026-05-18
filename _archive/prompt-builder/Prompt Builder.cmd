@echo off
cd /d "%~dp0"
node scripts/generate_prompt_builder_data.js
start "" "Prompt Builder.html"
