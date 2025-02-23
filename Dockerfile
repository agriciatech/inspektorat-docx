FROM ubuntu:latest
FROM node:20-alpine 

RUN apk update && apk add openssl

WORKDIR /usr/src/inspektorat-docx

COPY package.json /usr/src/inspektorat-docx/package.json
COPY package-lock.json /usr/src/inspektorat-docx/package-lock.json

RUN npm install
COPY . /usr/src/inspektorat-docx

CMD ["npm", "start"]