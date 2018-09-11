
# 2018年同人誌 / SIer向けExcel VBA Tips用ソースコード

## Jenkins (マスター) の立ち上げ方

Docker + docker-compose が必要となります。

```
$ cd jenkins
$ docker-compose up

1. コンソール中に出るpasswordをコピーして起動画面 http://localhost:8080 にて貼り付けてください。
2. admin / admin でログインできるようにしておいてください。
```

## Jenkins (スレーブ) の立ち上げ方

1. `build.gradle` の中のJenkinsのURLなど適宜書き換えてください。
2. JDKインストール済みのWindowsマシンより、下記コマンドを実行すればスレーブとして繋がります。 (Swarm Plugin経由)

```
$ gradlew.bat
```

## License

MIT license.
