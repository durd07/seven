HUB=felixdu.dynamic.nsn-net.net
REPO=seven
TAG=latest

docker build --build-arg https_proxy=http://10.158.100.9:8080 -t $HUB/$REPO:$TAG .
docker push $HUB/$REPO:$TAG

docker run -td --restart=always -p 8501:8501 -v $(pwd):/app $HUB/$REPO:$TAG
