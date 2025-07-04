# GitWorkflow for SendToTTS

This project follows GitWorkflow for branch management and development process.

## Branch Structure

### Main Branches
- **master**: Production-ready, stable releases only
- **develop**: Main development branch (default branch)

### Supporting Branches
- **feature/**: New features and enhancements
- **release/**: Release preparation
- **hotfix/**: Critical fixes for production

## Workflow Rules

1. **All development** happens on `develop` branch or feature branches
2. **Feature branches** are created from `develop` and merged back to `develop`
3. **Release branches** are created from `develop` when ready for release
4. **Hotfix branches** are created from `master` for critical production fixes
5. **Master branch** only receives merges from release or hotfix branches

## Commands

### Start new feature
```bash
git checkout develop
git pull origin develop
git checkout -b feature/feature-name
```

### Finish feature
```bash
git checkout develop
git pull origin develop
git checkout feature/feature-name
git rebase develop
git checkout develop
git merge --no-ff feature/feature-name
git push origin develop
git branch -d feature/feature-name
```

### Create release
```bash
git checkout develop
git pull origin develop
git checkout -b release/v1.0.0
# Bump version, update changelog
git commit -m "Bump version to 1.0.0"
```

### Finish release
```bash
git checkout master
git merge --no-ff release/v1.0.0
git tag -a v1.0.0 -m "Release version 1.0.0"
git checkout develop
git merge --no-ff release/v1.0.0
git branch -d release/v1.0.0
git push origin master develop --tags
```

### Hotfix
```bash
git checkout master
git pull origin master
git checkout -b hotfix/v1.0.1
# Fix the issue
git commit -m "Fix critical issue"
git checkout master
git merge --no-ff hotfix/v1.0.1
git tag -a v1.0.1 -m "Hotfix version 1.0.1"
git checkout develop
git merge --no-ff hotfix/v1.0.1
git branch -d hotfix/v1.0.1
git push origin master develop --tags
```

## Repository Setup

- **Private repository**: ✅
- **MIT License**: ✅
- **Default branch**: `develop` ✅
- **GitWorkflow enabled**: ✅

## Branch Protection

Consider setting up branch protection rules:
- Protect `master` branch: require pull requests, require status checks
- Protect `develop` branch: require pull requests for external contributors 