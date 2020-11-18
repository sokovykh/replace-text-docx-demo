# Manual testing
docker build --pull -t replace-text-docx-demo .
docker run --rm -it -p 8000:80 replace-text-docx-demo:latest


# HEROKU PUBLISH
docker tag replace-text-docx-demo:latest registry.heroku.com/replace-text-docx-demo/web
heroku login
heroku container:login
heroku container:push web -a replace-text-docx-demo
heroku container:release web -a replace-text-docx-demo