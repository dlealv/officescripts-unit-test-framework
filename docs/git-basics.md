# Git Basic Operations Reference

## 1. Setup

```bash
git config --global user.name "Your Name"
git config --global user.email "you@example.com"
```

## 2. Creating a Repository

```bash
git init                # Initialize new repo in current directory
git clone <url>         # Clone an existing repository
```

## 3. Checking Status

```bash
git status              # Show status of changes
```

## 4. Staging and Committing
- **Stagging**: Moves changes from your working directory to the staging area, preparing them for the next commit
- **Committing**: Captures a snapshot of the currently staged changes and records them in the local repository
```bash
git add <file>                              # Stage a specific file
git add .                                   # Stage all files
git commit -m "Message"                     # Commit staged changes
git commit --allow-empty -m "Message"       # A commit with no changes
```

## 5. Viewing History

```bash
git log                 # View commit history
git log --oneline       # Condensed log
```

## 6. Working with Branches

```bash
git branch              # List branches
git branch <name>       # Create new branch
git checkout <name>     # Switch to branch
git checkout -b <name>  # Create and switch to branch
git merge <name>        # Merge branch into current
git branch -d <name>    # Delete a branch
```

## 7. Pulling and Pushing
- **Pulling**: Fetches changes from a remote repository and integrates them into your local branch
- **Pushing**: Upload local repository content to a remote repository
```bash
git pull                # Fetch and merge from remote
git push                # Push changes to remote
git push -u origin <branch> # Push new branch and track
```

## 8. Undoing Changes

```bash
git checkout -- <file>        # Discard changes in working directory
git reset HEAD <file>         # Unstage a file
git revert <commit>           # Create a new commit to undo changes
git reset --hard <commit>     # Reset history and working directory (danger!)
```

## 9. Tags

```bash
git tag                  # List tags
git tag <name>           # Create tag
git push origin <tag>    # Push tag to remote
```

## 10. Rebasing (Advanced)
Modify the commit history of a branch by moving or combining a sequence of commits to a new base commit
```bash
git rebase <base-branch>        # Rebase current branch onto base
git rebase -i <commit-hash>     # Interactive rebase for editing history
```

## 11. Stashing
Allows temporarily save uncommitted changes
```bash
git stash                  # Stash unsaved changes
git stash pop              # Apply stashed changes
```

---

**Tip:**  
Use `git help <command>` for detailed info about any command!
