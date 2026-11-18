/**
 * ダッシュボードのautoRefreshは元の更新処理だけに戻す。
 * フォーム受信・30分データ保存は、別プロジェクトのcall_statistics.gsが担当する。
 * 同じフォームを複数のGASから読み取って削除しない。
 */
function autoRefresh() {
  autoRefreshBase_();
  try { rebuildClosedMonths_(); } catch (e) { console.log(e); }
  aclRefresh_();
}
