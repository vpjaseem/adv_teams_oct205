#!/usr/bin/env bash
sudo apt update -y
sudo apt install -y apache2 php libapache2-mod-php unzip wget

sudo systemctl enable --now apache2

sudo chmod -R 777 /var/www/html

cd /tmp
wget -O main.zip "https://github.com/vpjaseem/az-lin-php-web-app/archive/refs/heads/main.zip"
unzip -o main.zip
sudo mv az-lin-php-web-app-main/* /var/www/html/
rm -rf az-lin-php-web-app-main main.zip

cd /var/www/html/
rm -r index.html

sudo systemctl restart apache2
