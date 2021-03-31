# FROM nginx:alpine
FROM nginx:alpine

LABEL description="Docker file for Test"

RUN mkdir -p /var/www/outlook

ADD  dist/outlookHello/ /var/www/outlook/public/
COPY nginx.conf /etc/nginx/nginx.conf
