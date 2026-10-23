# Developement

## Enable github pages to serve index.html

From github CLI Create a new repo and enable pages
```
gh repo create bjacques/40k-rpg --public --source=. --remote=origin --push
gh api -X POST repos/bjacques/40k-rpg/pages -f "source[branch]=main" -f "source[path]=/" --jq '.html_url'
```
