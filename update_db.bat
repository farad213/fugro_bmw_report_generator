@echo off

set "git_repo_path=.git"
set "file_to_push=source/database/BMW_report_buildings_data.xlsx"

cd /d "%git_repo_path%"

git add "%file_to_push%"
git commit -m "Added %file_to_push%"
git push

echo File pushed to Git repository.

pause