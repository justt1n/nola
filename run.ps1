echo "Running Script"
echo "Checking for update"
git pull
echo "Update completed"
& myenv\Scripts\Activate.ps1
echo "Running Script..."
python script.py
echo "Complete"
Read-Host -Prompt "Press Enter to exit"