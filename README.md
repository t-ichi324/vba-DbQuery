# vba-DbQuery
-VBA (Microsoft.Access)のデータベースアクセス用クラスモジュールです。
-プロジェクトのIDEへ「DbQuery.cls」と「DbReader.cls」をインポートしてください。

## [USER]テーブルの [ID] = 1 の[NAME]を取得
```
Dim q As New DbQuery
Debug.Print q.Table("USER").Wheres("ID", 1).GetField("NAME")
```

## [USER]テーブルの [ID] > 1 のレコードセットを取得
```
Dim q As New DbQuery
Dim rs As Recordset
Set rs = q.Table("USER").Wheres("ID", 1, ">").GetRecordset
```

## [USER]テーブルの [NAME] LIKE '+Tarou*' のDbReaderオブジェクトを取得
```
Dim q As New DbQuery
Dim rdr As DbReader
Set rdr = q.Table("USER").Wheres("ID", "*TAROU*", "LIKE").GetReader
Do While rd.ReadLine
    Debug.Print rd.GetStr("NAME")
Loop
```

## [USER]テーブルにユーザを登録
```
Dim q As New DbQuery
Call q.Table("USER").Sets("ID", 10).Sets("NAME", "TEST-10").Insert
Call q.Table("USER").Sets("ID", 11).Sets("NAME", "TEST-11").Insert
Call q.Table("USER").Sets("ID", 12).Sets("NAME", "TEST-12").Insert
```

## [USER]テーブルの[ID] = 11 OR [ID] = 12の[NAME]を更新
```
Dim q As New DbQuery
Call q.Table("USER").Sets("NAME", "updated").Wheres("ID", 11).Ors("ID", 12).Update
```

## [USER]テーブルの [ID] = 12 を削除
```
Dim q As New DbQuery
Call q.Table("USER").Wheres("ID", 12).Delete
```

