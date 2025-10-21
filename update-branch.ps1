param(
    [string]$sourceBranch,   # branch to copy from
    [string]$sourceFolder,   # folder in the source branch
    [string]$targetBranch,   # branch to copy into
    [string]$targetFolder    # folder in the target branch
)

# Switch to the target branch
git checkout $targetBranch

# Checkout the folder from the source branch
git checkout $sourceBranch -- $sourceFolder

# Copy the contents to the target folder
Copy-Item -Path "$sourceFolder\*" -Destination "$targetFolder" -Recurse -Force

# Add changes to git
git add $targetFolder

# Create a commit
git commit -m "Updated $targetFolder from $sourceBranch"

# Push to remote
git push
