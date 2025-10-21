param(
    [string]$sourceBranch,  # branch to copy from (e.g., develop, test)
    [string]$targetBranch   # branch to copy into (e.g., test, prod)
)

# Provjeri postoji li target branch lokalno, ako ne kreiraj ga
$branches = git branch --list $targetBranch
if (-not $branches) {
    Write-Host "Target branch '$targetBranch' does not exist locally. Creating it..."
    git checkout -b $targetBranch
    git push -u origin $targetBranch
} else {
    git checkout $targetBranch
}

# Merge all changes from source branch into target branch
git merge $sourceBranch --no-ff -m "Merged all changes from $sourceBranch into $targetBranch"

# Push changes to remote
git push

Write-Host "✅ Successfully merged '$sourceBranch' into '$targetBranch' and pushed all changes!"
