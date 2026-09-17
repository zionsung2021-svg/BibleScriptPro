FROM nginx:alpine
# 내 컴퓨터에서 이미 빌드된 dist 폴더를 Nginx 서버로 바로 복사합니다.
COPY dist /usr/share/nginx/html
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]