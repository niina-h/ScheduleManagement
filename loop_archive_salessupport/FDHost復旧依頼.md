# FDHost サービス復旧依頼(DBA/インフラ担当宛)

- 作成日:2026-07-21(反復 #022)
- 対象サーバー:`01-DB-SVR-01\SQLSTANDARD`
- 対象 DB:`SalesSupportDB_Integration`
- 依頼元:営業支援ツール移設プロジェクト(タスク 3-9:全文検索(FTS)導入)

---

## 事象

`SalesSupportDB_Integration` にて `T_SalesWork` テーブルへ全文検索(FTS)インデックスを作成した。
カタログ・サロゲートキー列・一意インデックス・FTS インデックスの作成はすべて成功したが、
実際に `CONTAINS()` 関数を使った検索を実行すると以下のエラーが発生する。

```sql
SELECT COUNT(*) FROM T_SalesWork WHERE CONTAINS(SalesNaiyou, N'キーワード');
```

```
フルテキスト フィルター デーモン ホスト(FDHost)プロセスとの通信中に SQL Server でエラー
0x8007042d が発生しました。FDHost プロセスが実行中であることを確認してください。
FDHost プロセスを再開するには、sp_fulltext_service 'restart_all_fdhosts' コマンドを実行するか、
SQL Server インスタンスを再起動してください。
```

`FULLTEXTCATALOGPROPERTY('SalesSupportFTS', 'ItemCount')` も `0` のままで、
母集団化(クロール)自体が進んでいない状態と一致している。

## 依頼したい対応(優先順)

1. **Windows サービス `SQL Full-text Filter Daemon Launcher (MSSQLSERVER)`(または該当インスタンス名)の状態確認・再起動**
   - サーバー `01-DB-SVR-01` にログインし、サービスマネージャーまたは PowerShell で状態確認
   - `Get-Service -Name "*FDLauncher*"` 等で対象サービスを特定し、停止していれば起動
2. 上記で解決しない場合、**`sysadmin` 権限を持つアカウントで以下を実行**
   ```sql
   EXEC sp_fulltext_service 'restart_all_fdhosts';
   ```
   (現在使用しているアプリケーション接続アカウント `CrownSSDAdmin` には `sysadmin` 権限がなく実行不可のため確認できていない)
3. それでも解決しない場合、**SQL Server インスタンスの再起動**
   - ⚠️ **このインスタンスには 20 個の本番データベースが同居している**(`SalesSupportDB` 本番含む、`CrownAIDB`/`CrownGateDB`/`CrownHaikinDB`/`CrownKaisekiDB`/`KintaiManageDB`/`ElectronicApplicationDB` 等)
   - インスタンス再起動は**全システムに影響**するため、必ず事前に関係者への周知・保守時間帯の調整をお願いしたい

## 確認後にお願いしたいこと

FDHost 復旧後、以下を実行して動作確認いただけると助かる:

```sql
USE SalesSupportDB_Integration;
SELECT FULLTEXTCATALOGPROPERTY('SalesSupportFTS', 'PopulateStatus') AS PopulateStatus,
       FULLTEXTCATALOGPROPERTY('SalesSupportFTS', 'ItemCount') AS ItemCount;
-- ItemCount が 75000 件超になれば母集団化完了

SELECT COUNT(*) FROM T_SalesWork WHERE CONTAINS(SalesNaiyou, N'樋門');
-- エラーなく件数が返れば FTS 機能は正常動作
```

## 参考:今回作成したオブジェクト(すでに適用済み)

- カタログ:`SalesSupportFTS`
- 追加列:`T_SalesWork.FTSKeyID`(INT IDENTITY、サロゲートキー)
- 一意インデックス:`UQ_T_SalesWork_FTSKeyID`
- FTS インデックス:`T_SalesWork(Mokuteki, SalesNaiyou)` LANGUAGE 1041(日本語)、CHANGE_TRACKING AUTO

スクリプト本体:`SQL/v0.6_統合版/02_StoredProcedures/14_FTS_T_SalesWork.sql`
ロールバック手順:`SQL/v0.6_統合版/99_Rollback/00_Rollback_All.sql`

---

## 変更履歴

| 日付 | 内容 |
|---|---|
| 2026-07-21 | 初版作成(反復 #022) |
