// 00_README.md.gs
/**
 * # interface
 *
 * 個別の生徒の年間計画レビューと、この「年間計画レビューテンプレート」ライブラリとの接続部分のふるまいを定義します。
 *
 *
 * 以下のGAS参照用ガントチャートの実装と意義はほぼ同じで、現状年間計画レビューに必要な機能のみをこちらに移植した状態です。
 * 詳しい説明は以下のurlの00番のドキュメントをご覧ください。
 *
 * @link {https://script.google.com/u/0/home/projects/1QtuU9ITUCuHU4ZCts5_WLGUC8aF5l9vg8Avpbf7WDMDvPjqHk-XCxl3R/edit}
 *
 *
 * サイドバーから任意の関数を実行できる機能については現状移植していません。
 *
 *
 * 個別の生徒の呼び出し用GAS実装
 * -------------------------------------------------------------------------------------
 * //共通ライブラリに処理の中身の全てを投げる。abstractFunctionは、任意の関数を実行するための抽象ラッパー
 * function onOpen(){ return YearlyPlanReviewLib.onOpenAction();}
 * function onEdit(e){ return YearlyPlanReviewLib.onEditAction(e);}
 * function onInstall(e){ return YearlyPlanReviewLib.onInstallAction(e);}
 * function onChange(e){ return YearlyPlanReviewLib.onChangeAction(e);}
 * function onSelectionChange(e){ return YearlyPlanReviewLib.onSelectionChangeAction(e);}
 * function onFormSubmit(e){ return YearlyPlanReviewLib.onFormSubmitAction(e);}
 * function doGet(e){ return YearlyPlanReviewLib.onGetAction(e);}
 * function doPost(e){ return YearlyPlanReviewLib.onPostAction(e);}
 * function abstractFunction(request) { return YearlyPlanReviewLib.abstractFunction(request); }
 * -------------------------------------------------------------------------------------
 *
 */
