# -*- coding: utf-8 -*-
"""
create_excel_guide.py
Generates the comprehensive JTA Transaction Timeout & 2PC Guide Excel workbook
formatted with 'Meiryo UI' font and professional cell styling.
"""

import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_workbook():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # -------------------------------------------------------------
    # Styling Palette & Helpers
    # -------------------------------------------------------------
    FONT_FAMILY = "Meiryo UI"
    
    # Fonts
    font_title = Font(name=FONT_FAMILY, size=15, bold=True, color="FFFFFF")
    font_section = Font(name=FONT_FAMILY, size=11, bold=True, color="1B365D")
    font_header = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    font_sub_header = Font(name=FONT_FAMILY, size=9.5, bold=True, color="1B365D")
    font_body = Font(name=FONT_FAMILY, size=9.5, color="222222")
    font_body_bold = Font(name=FONT_FAMILY, size=9.5, bold=True, color="222222")
    font_body_highlight = Font(name=FONT_FAMILY, size=9.5, bold=True, color="C00000")
    font_code = Font(name="Consolas", size=9, color="1B365D")
    font_note = Font(name=FONT_FAMILY, size=8.5, italic=True, color="555555")

    # Fills
    fill_title = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # Navy
    fill_header = PatternFill(start_color="2A5298", end_color="2A5298", fill_type="solid") # Steel Navy
    fill_sub_header = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") # Light Ice Blue
    fill_zebra_even = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_zebra_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_alert = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Soft Orange / Alert
    fill_success = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Soft Green
    fill_accent = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")

    # Borders
    thin_border_side = Side(border_style="thin", color="BDC3C7")
    thick_bottom_side = Side(border_style="medium", color="1B365D")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom_side)

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    def apply_title_row(ws, title_text, max_col):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        cell = ws.cell(row=1, column=1, value=title_text)
        cell.font = font_title
        cell.fill = fill_title
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 40

    def apply_section_row(ws, row_idx, section_text, max_col):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=max_col)
        cell = ws.cell(row=row_idx, column=1, value=section_text)
        cell.font = font_section
        cell.fill = fill_sub_header
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row_idx].height = 26

    def style_table_headers(ws, row_idx, headers):
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_header
        ws.row_dimensions[row_idx].height = 28

    def adjust_column_widths(ws, max_col, custom_widths=None):
        ws.views.sheetView[0].showGridLines = True
        for col in range(1, max_col + 1):
            col_letter = get_column_letter(col)
            if custom_widths and col in custom_widths:
                ws.column_dimensions[col_letter].width = custom_widths[col]
            else:
                max_len = 0
                for cell in ws[col_letter]:
                    val = str(cell.value or "")
                    # simple heuristic for jp text
                    line_len = max([len(line.encode('shift-jis', errors='ignore')) if line else 0 for line in val.split("\n")] or [0])
                    if line_len > max_len:
                        max_len = line_len
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    # =========================================================================
    # SHEET 1: 小学生向け・ストーリー解説 (Analogy & Story)
    # =========================================================================
    ws1 = wb.create_sheet(title="1_小学生向け解説とたとえ話")
    apply_title_row(ws1, "【小学生でも絶対にわかる！】JTAトランザクションとタイムアウトの秘密（超巨大ケーキ作り大作戦）", 6)
    
    row = 3
    apply_section_row(ws1, row, "1. 登場人物と役割の対応表（これさえ読めば全部つながる！）", 6)
    row += 1
    
    headers_s1_char = ["現実の世界（ITシステム）", "たとえ話の世界（学園祭のケーキ作り）", "担当する仕事・役割", "もしこれがないとどうなるか？（事故例）"]
    # We will use columns A to D, merge D & E or use 5 columns
    headers_s1 = ["専門用語（IT用語）", "たとえ話のキャラクター", "役割・仕事内容", "普段の動き（通常時）", "トラブル時の動き（制限時間オーバー時）"]
    style_table_headers(ws1, row, headers_s1)
    row += 1

    char_data = [
        ("Narayana\n(トランザクション管理者)", "頼れる生徒会長\n（全体の指揮係）", "みんなに声をかけ、作業の開始から最後の合体までを取り仕切る", "全員に「準備できた？」と聞き、全員「できた！」となったら「せーの、合体！」と合図する", "誰かが倒れたら「作戦中止！みんなケーキを片付けて元に戻して！」とホイッスルを鳴らす"),
        ("2PC (2フェーズコミット)", "合体直前の「二段階確認」\n（準備確認 ➔ 一斉合体）", "2つの係（スポンジ係・生クリーム係）が同時にケーキを完成させる約束事", "①「準備OK？」(Phase 1: Prepare)\n②全員OKなら「合体！」(Phase 2: Commit)", "片方でも「失敗した！」と言ったら、もう片方も塗ったクリームを捨てて元通りにする"),
        ("JTAトランザクション\nタイムアウト", "学園祭開始までの\n【全体リミット（30分）】", "ケーキ作りプロジェクト全体の制限時間。これを超えたら学園祭が始まってしまう！", "30分以内にケーキが完成すれば大成功！お客様に笑顔で提供する", "30分を過ぎたら、生焼けのケーキを出すわけにいかないので、全作業を強制中止（全ロールバック）"),
        ("クエリタイムアウト\n(Query Timeout)", "イチゴを切る1作業の\n【個別リミット（3分）】", "包丁でイチゴを刻む「たった1つの作業」の制限時間", "サクサクと2分でイチゴを刻み終わって次の作業へ進む", "包丁を持ったまま居眠りして3分経ったら、タイマーが鳴って「イチゴ切りは一旦やめ！」と中断する"),
        ("set-tx-query-timeout\n(連動設定)", "残り時間に自動で合わせる\n【賢いスマートウォッチ】", "全体の残り時間を見て、個別作業のタイマーを自動で短縮してくれる便利機能", "全体が残り25分あれば、イチゴ切りは普通に3分制限で動く", "全体が残り1分しかない時、イチゴ切りタイマーも自動で「1分」に縮める（無駄な作業をさせない）"),
        ("xa-resource-timeout\n(XAリソース制限時間)", "学校のオーブン使用予約の\n【占有リミット】", "調理室のオーブン（データベース）に「最長でもこれだけしか使わないよ」と宣言しておく時間", "予定通り10分でケーキを焼き上げて、オーブンを綺麗に空ける", "もし生徒会長が倒れて連絡が途絶えても、オーブン側が時間切れで自動的にスイッチを切りロック解除する"),
    ]

    for item in char_data:
        fill = fill_zebra_even if row % 2 == 0 else fill_zebra_odd
        for c_idx, val in enumerate(item, start=1):
            cell = ws1.cell(row=row, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = align_left if c_idx > 2 else align_center
        ws1.row_dimensions[row].height = 42
        row += 1

    row += 2
    apply_section_row(ws1, row, "2. なんで「2PCコミット」があるのに「JTAタイムアウト」が必要なの？（核心の疑問）", 5)
    row += 1

    q_and_a = [
        ("素朴な疑問：", "「生徒会長が『準備できた？』➔『せーの、合体！』って2段階で確認（2PC）してくれるなら、失敗したらやり直せるんだから、全体リミット（JTAタイムアウト）なんて決めなくても大丈夫なんじゃないの？」"),
        ("答え（大ピンチ！）：", "【絶対にダメです！むしろ2つの係が協力するからこそ、全体リミットがないと学校中が大パニックになります！】"),
        ("理由①：『巻き添えフリーズ』", "スポンジ係が「イチゴが固くて切れない…」と包丁を持ったままフリーズしたとします。\nすると、生クリーム係は「スポンジが来ないとクリームを塗れないよ…」と泡立て器を持ったまま一生待たされます！\n生クリームは溶け、調理室の水道も占有され、他のクラスの生徒も調理室を使えなくなります（これが『データベースのロック独占』です）。"),
        ("理由②：『生徒会長が声をかける前に力尽きる』", "2PC（二段階確認）は、ケーキの材料がぜんぶ揃って「さあ仕上げだ！」という段階になって初めて始まります。\n材料を切っている途中（Phase 0: 業務処理中）に誰かが居眠りしたら、生徒会長はそもそも「準備できた？」と声をかけることすらできません！\n全体タイマー（JTAタイムアウト）が鳴って生徒会長が「時間切れ！解散！」と叫ばない限り、全員が石のように固まり続けます。"),
        ("理由③：『無駄な居残り作業の防止』", "学園祭開始まであと10秒しかないのに、スポンジ係が「これから手の込んだ飾り付けを始めます（所要時間5分）」と作業を始めたらどうでしょう？\n完成した頃には学園祭は終わっています。全体の残り時間（JTA）を意識しないと、絶対に間に合わない無駄な作業に時間と体力を浪費してしまいます。"),
    ]

    for title, desc in q_and_a:
        ws1.cell(row=row, column=1, value=title).font = font_body_bold
        ws1.cell(row=row, column=1).alignment = align_left
        ws1.cell(row=row, column=1).fill = fill_accent
        ws1.cell(row=row, column=1).border = border_cell

        ws1.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
        for c in range(2, 6):
            cell = ws1.cell(row=row, column=c)
            cell.border = border_cell
            cell.fill = fill_zebra_odd
        top_cell = ws1.cell(row=row, column=2, value=desc)
        top_cell.font = font_body_highlight if "大ピンチ" in desc else font_body
        top_cell.alignment = align_left_top
        ws1.row_dimensions[row].height = 46 if len(desc) > 80 else 28
        row += 1

    adjust_column_widths(ws1, 5, {1: 22, 2: 24, 3: 32, 4: 36, 5: 38})

    # =========================================================================
    # SHEET 2: 4大設定値の徹底比較 (Deep Comparison)
    # =========================================================================
    ws2 = wb.create_sheet(title="2_4大タイムアウト徹底比較")
    apply_title_row(ws2, "【完全網羅】JTAタイムアウト / xa-resume(resource)-timeout / クエリタイムアウト / set-tx-query-timeout 徹底比較", 8)

    row = 3
    apply_section_row(ws2, row, "1. 4大タイムアウト設定値 詳細スペック比較マトリクス", 8)
    row += 1

    headers_s2 = [
        "設定項目名", "正式設定パラメータ・場所", "管理レイヤー\n(制御主体)", "制御対象\n(スコープ)",
        "動作原理と仕組み\n(裏側で何が起きるか)", "登場の歴史と背景\n(なぜ必要になったか？)",
        "デフォルト値と\n推奨設定値", "タイムアウト発生時の挙動\nとスローされる例外"
    ]
    style_table_headers(ws2, row, headers_s2)
    row += 1

    s2_data = [
        (
            "JTAトランザクション\nタイムアウト\n(Global Tx Timeout)",
            "【XML】transactionsサブシステム\n<coordinator-environment\n  default-timeout=\"300\"/>\n【CLI】/subsystem=transactions:\nwrite-attribute(name=default-timeout, value=300)\n【Java】@TransactionTimeout(300)\nUserTransaction.setTransactionTimeout(300)",
            "JBoss EAP\n(Narayana TM\nトランザクションマネージャ)",
            "トランザクション全体の寿命\n(UserTransaction.beginから\ncommit/rollback完了まで)\n※複数DB・JMS・EJBにまたがる\nグローバルスコープ",
            "Narayanaの「TransactionReaper」スレッドが管理。\n①トランザクション開始時にReaperリストへ登録。\n②タイマー満了時、ReaperWorkerThreadがTxステータスをMARKED_ROLLBACKに変更。\n③未コミットのリソースに対しXAResource.rollback()を実行して強制アボート。",
            "【背景・歴史】\n1990年代、分散トランザクション(X/Open DTP)が普及した際、単一SQLのタイムアウトだけでは「複数DBにまたがる長時間のロック保持」や「無限ループによるリソース枯渇」を防げず、システム全体が連鎖停止する大事故が多発。\nそのため、個別のSQLに関係なく『業務処理全体の最大許容時間』を強制打ち切りするマスター監視役として誕生した。",
            "【デフォルト値】\n300秒 (5分)\n\n【本番推奨値】\nオンラインWeb系: 30〜60秒\nバッチ・夜間処理: 300〜1800秒\n(※業務SLA・呼出元タイムアウトより必ず短く設定)",
            "・トランザクションがロールバック専用状態になる。\n・コミット試行時に javax.transaction.RollbackException をスロー。\n・ログに ARJUNA016029 / ARJUNA012117 (Transaction has timed out) が出力される。"
        ),
        (
            "xa-resource-timeout\n(※xa-resume-timeout\nの文脈含む)",
            "【XML】datasourcesサブシステム\n<xa-datasource ...>\n  <timeout>\n    <xa-resource-timeout>300</xa-resource-timeout>\n  </timeout>\n</xa-datasource>\n【CLI】/subsystem=datasources/\nxa-data-source=XXX:\nwrite-attribute(name=xa-resource-timeout, value=300)",
            "JCAコンテナ (IronJacamar)\n➔ 各リソースマネージャ\n(Oracle/PostgreSQL等のDB)",
            "DB側のXAトランザクション\nブランチ寿命\n(XAResource単位)\n※DBサーバ内部のロック保持時間",
            "JCAがコネクション接続時、JDBCのXAResource.setTransactionTimeout(秒)を呼び出してDBエンジンに伝達。\nDB側内部タイマーがセットされ、指定時間内にprepare/commitが届かない場合、DBが自主的にブランチを破棄・ロールバックする。\n※Oracle等で別セッションのサスペンド待ちをする「resume timeout(ORA-25351)」の制御とも密接に関連。",
            "【背景・歴史】\nJBoss側のJTAタイムアウトだけでは、アプリサーバが突然クラッシュやNW切断した際、DB側に未解決のXAブランチ(トランザクションロック)が半永久的に残り続ける問題が生じた。\n「DB自身にもタイムアウト値を教えて、自律的に不要なロックを解放させたい」という要求から、JTA仕様のXAResource.setTransactionTimeoutに対応するJCA設定として生まれた。",
            "【デフォルト値】\n0 (無効/DBのデフォルト)\n\n【本番推奨値】\nJTAタイムアウトと同値、または\nJTAタイムアウトより若干長め(+5〜10秒)\n※短くするとヒューリスティック例外を誘発！",
            "・DB側が自律ロールバックした場合、Narayanaがprepare/commitしようとすると XAException (XAER_NOTA: 未知のXID) が発生。\n・最悪の場合、一部だけコミットされるヒューリスティックハザード (XA_HEURHAZ) に発展する。"
        ),
        (
            "クエリタイムアウト\n(Query Timeout)",
            "【XML】datasourcesサブシステム\n<datasource/xa-datasource ...>\n  <timeout>\n    <query-timeout>30</query-timeout>\n  </timeout>\n【Java】\njava.sql.Statement.setQueryTimeout(30)\nJPA: javax.persistence.query.timeout",
            "JDBCドライバ\n➔ データベース",
            "単一のSQL文の実行時間\n(executeQuery / executeUpdate)\n※1回ごとのSQL発行単位",
            "JDBCドライバ内部のTimerスレッドが監視。\nSQL実行開始と同時にタイマー起動。指定秒数を超えても結果が戻らない場合、ドライバが別ソケットまたはOOB(Out-of-Band)パケットでDBへ「SQL_CANCEL」割込みを送信。\nDBが実行中のSQL実行計画を中断し、ドライバへキャンセル応答を返す。",
            "【背景・歴史】\nRDBMS初期から存在する最も基本的な防壁。\nインデックス未適用によるフルスキャン、テーブルデッドロック、複雑なJOINによる暴走クエリがDBのCPU/I/Oを100%占有し、後続の全トランザクションを道連れにするのを防ぐために誕生した。",
            "【デフォルト値】\n0 (無制限/ドライバ依存)\n\n【本番推奨値】\n一般的なSQL: 10〜30秒\n(※JTAトランザクションタイムアウトよりも必ず小さく設定する)",
            "・実行中スレッドに java.sql.SQLTimeoutException がスローされる。\n・該当SQLは即座に中断されるが、JTAトランザクション自体は例外をキャッチして適切に処理可能（ロールバックするか判断できる）。"
        ),
        (
            "set-tx-query-timeout\n(動的クエリ連動)",
            "【XML】datasourcesサブシステム\n<datasource/xa-datasource ...>\n  <timeout>\n    <set-tx-query-timeout>true</set-tx-query-timeout>\n    <query-timeout>30</query-timeout>\n  </timeout>\n【CLI】/subsystem=datasources/\nxa-data-source=XXX:\nwrite-attribute(name=set-tx-query-timeout, value=true)",
            "JBoss EAP\n(IronJacamar JCAラッパー)",
            "JTAトランザクション内で発行\nされるすべてのSQL文\n(動的にクエリ制限時間を計算)",
            "IronJacamarのWrappedConnectionがStatement/PreparedStatement生成時に割り込み。\n①現在スレッドにアクティブなJTA Txが存在するか確認。\n②存在する場合、JTAの残り時間 (txTimeout - 経過時間) を算出。\n③その残り秒数を Statement.setQueryTimeout(残り秒数) に動的設定！\n※Tx外なら静的なquery-timeoutを使用。",
            "【背景・歴史】\n「JTAタイムアウトが残り3秒しかないのに、30秒の静的クエリタイムアウトで重いSQLを投げてしまい、DBが無駄に20秒処理した挙句コミット時にJTA時間切れで全部やり直しになる」という『無駄打ち問題』が頻発。\n「トランザクションの余命に合わせて、自動でクエリ制限時間をカウントダウンして注入できないか？」という発想からJBoss/IronJacamarに実装された。",
            "【デフォルト値】\nfalse (無効)\n\n【本番推奨値】\ntrue (強く推奨！)\n※ただし古いJDBCドライバではsetQueryTimeoutが非同期で不安定な場合があるため実機検証必須",
            "・JTAの残り時間を超えそうになった段階で、DBが無駄なSQLを即座にキャンセル (SQLTimeoutException)。\n・無駄なDBリソース消費を防ぎ、即座にJTAタイムアウト処理へ移行できる。"
        )
    ]

    for item in s2_data:
        fill = fill_zebra_even if row % 2 == 0 else fill_zebra_odd
        for c_idx, val in enumerate(item, start=1):
            cell = ws2.cell(row=row, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = align_left_top if c_idx in [2, 4, 5, 6, 7, 8] else align_center
        ws2.row_dimensions[row].height = 120
        row += 1

    adjust_column_widths(ws2, 8, {1: 20, 2: 30, 3: 18, 4: 22, 5: 38, 6: 38, 7: 28, 8: 34})

    # =========================================================================
    # SHEET 3: Narayana 2PCとJTA必要性検討 (Deep Architectural Analysis)
    # =========================================================================
    ws3 = wb.create_sheet(title="3_Narayana2PCとJTA必要性")
    apply_title_row(ws3, "【技術深堀り】JBoss EAP Narayana 2PCコミットにおけるJTAトランザクションタイムアウトの必要性徹底検証", 6)

    row = 3
    apply_section_row(ws3, row, "1. 核心命題：「Narayanaで2PCコミットを使っているなら、JTAタイムアウトは不要か？」に対する絶対的結論", 6)
    row += 1

    conclusion_box = [
        ("結論：", "【極めて重要であり、絶対に必要不可欠（むしろ2PCを採用しているからこそ、最重要の安全装置）】"),
        ("理由の要約：", "2PCプロトコルは『全リソースが準備完了した後のコミット整合性』を保証する仕組みであり、『準備前のフリーズ』や『ロック爆発』を自律解除する能力を持たないため。")
    ]
    for k, v in conclusion_box:
        ws3.cell(row=row, column=1, value=k).font = font_body_bold
        ws3.cell(row=row, column=1).fill = fill_alert
        ws3.cell(row=row, column=1).border = border_cell
        ws3.cell(row=row, column=1).alignment = align_center

        ws3.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        for c in range(2, 7):
            cell = ws3.cell(row=row, column=c)
            cell.border = border_cell
            cell.fill = fill_alert
        top_cell = ws3.cell(row=row, column=2, value=v)
        top_cell.font = font_body_highlight if "絶対" in v else font_body_bold
        top_cell.alignment = align_left
        ws3.row_dimensions[row].height = 30
        row += 1

    row += 2
    apply_section_row(ws3, row, "2. トランザクション・ライフサイクルの各フェーズとタイムアウトの挙動", 6)
    row += 1

    headers_s3_phase = ["フェーズ", "プロトコルの状態", "実行中の処理内容", "タイムアウト発生時の挙動", "2PC自体の動作", "JTAタイムアウトがない場合の危険度"]
    style_table_headers(ws3, row, headers_s3_phase)
    row += 1

    phases_data = [
        (
            "Phase 0\n業務処理実行中\n(Active状態)",
            "STATUS_ACTIVE\n(2PC未開始)",
            "・Java業務ロジックの実行\n・DB-AへのINSERT/UPDATE\n・DB-BへのUPDATE\n・JMSメッセージ送受信\n・外部REST/SOAP呼出",
            "NarayanaのTransactionReaperが検知。\nTxステータスをSTATUS_MARKED_ROLLBACKに強制変更。\n後続のSQL実行やcommit()呼出はすべて即座に拒否・アボートされる。",
            "2PCはまだ始まってすらいない。\n(2PCのプロトコルエンジンは待機中)",
            "【破滅的リスク：危険度 MAX】\nもしSQLや外部通信が無限待ちになった場合、永遠にトランザクションが終わらない。\nDB-AとDB-Bの両方で行ロック・コネクションが独占され、数分で全プールが枯渇して全システムが完全停止する！"
        ),
        (
            "Phase 1\n準備フェーズ\n(Prepare Phase)",
            "STATUS_PREPARING\n➔ STATUS_PREPARED",
            "Narayanaが全XAResourceに対して\nXAResource.prepare() を発行。\n全リソースがコミット可能か投票する。\n(全員OKならXA_OKを返却)",
            "Reaperまたはprepareタイムアウトが検知。\n参加リソースのうち1つでも応答がない、または時間切れの場合、Narayanaは全体をABORTと決定。\n全参加リソースにXAResource.rollback()を送信！",
            "2PCの第1段階。\n全員がXA_OKを返さない限り、絶対にコミットには進まない。",
            "【高リスク：危険度 大】\n参加DBの1つがNW障害で無応答になった場合、他の正常なDBはprepare完了状態でロックを抱えたまま待たされる（In-Doubt状態）。JTAタイムアウトでロールバックに倒す必要がある。"
        ),
        (
            "運命の分岐点\nコミットログ書き込み\n(Point of No Return)",
            "STATUS_COMMITTING\n(ディスク永続化)",
            "NarayanaがArjuna ObjectStore(ディスク)\nに「このTxはコミットする」という\nトランザクションログを物理書き込み完了！",
            "★この瞬間からタイムアウトによるロールバックは『絶対に不可』となる！\n以降はどんなに時間がかかっても、必ずコミットを完了させなければならない（整合性保護）。",
            "コミットの決定が確定。\nこの瞬間以降、Narayanaは絶対に後戻りしない。",
            "※ここは一瞬（ミリ秒オーダー）のディスクI/O処理。\nディスク故障時はRecoveryManagerが後で回復。"
        ),
        (
            "Phase 2\nコミット完了フェーズ\n(Commit Phase)",
            "STATUS_COMMITTING\n➔ STATUS_COMMITTED",
            "Narayanaが全XAResourceに対して\nXAResource.commit() を発行。\n変更を最終確定しロックを解放する。",
            "タイムアウトはロールバックを引き起こさない。\nNW切断等でcommit()が届かない場合、Narayanaの「Periodic Recovery」機能がバックグラウンドで再送し続け、確実にコミットを完遂する。",
            "2PCの第2段階。\n全リソースに確定指示を伝達する。",
            "【中リスク：ヒューリスティック例外】\nもしDB側のxa-resource-timeoutがJTAより短く、DBが勝手にロールバックしていた場合、「DB-Aはコミット、DB-Bはロールバック」というデータ不整合（ヒューリスティックハザード）が発生する！"
        )
    ]

    for item in phases_data:
        fill = fill_zebra_even if row % 2 == 0 else fill_zebra_odd
        for c_idx, val in enumerate(item, start=1):
            cell = ws3.cell(row=row, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = align_left_top if c_idx in [3, 4, 6] else align_center
        ws3.row_dimensions[row].height = 95
        row += 1

    row += 2
    apply_section_row(ws3, row, "3. Narayana内部構造：TransactionReaper（刈り取り屋）の動作メカニズム", 6)
    row += 1

    reaper_mechanisms = [
        ("監視方式 (Reaper Mode)", "DYNAMIC モード（推奨/標準） vs PERIODIC モード", "DYNAMICモードでは、登録されたトランザクションの最短タイムアウト時刻に合わせてスレッドが自律的に高精度スリープする。PERIODICモードは一定間隔(120秒等)で巡回する。EAPでは高負荷時も遅延なくタイムアウトを執行する。"),
        ("キャンセル実行 (Worker)", "ReaperWorkerThread の非同期処理", "タイマーが満了すると、Reaperは直接重い処理をせず、専用のワーカー(ReaperWorkerThread)にアボート処理を委譲する。これにより、監視ループ自体がブロックされるのを防止している。"),
        ("ゾンビ化防止パラメータ", "cancel-wait-period (500ms)\ncancel-fail-wait-period (500ms)", "キャンセル処理が500ms以内に終わらない場合、ワーカーにスレッド割込み(interrupt)をかける。それでも終了しない場合、トランザクションを「ゾンビ(WEDGED)」と判定し、警告ログ(ARJUNA012378)を出してシステム全体への巻き添えを防ぐ。"),
        ("DBソケット待ちとの闘い", "Javaスレッド強制停止不可問題", "JavaのThread.stop()は危険なため使用されない。もしスレッドがDBからのソケット受信(socketRead0)で完全ブロックしている場合、Reaperがステータスを変更してもスレッドは気づけない！➔ だからこそ『クエリタイムアウト』や『ソケットタイムアウト』が絶対にセットで必要なのだ！")
    ]

    for title, subtitle, detail in reaper_mechanisms:
        ws3.cell(row=row, column=1, value=title).font = font_body_bold
        ws3.cell(row=row, column=1).alignment = align_center
        ws3.cell(row=row, column=1).fill = fill_accent
        ws3.cell(row=row, column=1).border = border_cell

        ws3.cell(row=row, column=2, value=subtitle).font = font_body_bold
        ws3.cell(row=row, column=2).alignment = align_center
        ws3.cell(row=row, column=2).fill = fill_sub_header
        ws3.cell(row=row, column=2).border = border_cell

        ws3.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        for c in range(3, 7):
            cell = ws3.cell(row=row, column=c)
            cell.border = border_cell
            cell.fill = fill_zebra_odd
        top_cell = ws3.cell(row=row, column=3, value=detail)
        top_cell.font = font_body
        top_cell.alignment = align_left_top
        ws3.row_dimensions[row].height = 42
        row += 1

    adjust_column_widths(ws3, 6, {1: 22, 2: 24, 3: 32, 4: 34, 5: 24, 6: 38})

    # =========================================================================
    # SHEET 4: タイムアウト階層構造と整合性 (Hierarchy, Formulas & Best Practices)
    # =========================================================================
    ws4 = wb.create_sheet(title="4_タイムアウト階層構造と整合性")
    apply_title_row(ws4, "【設計ガイド】システム全体のタイムアウト階層構造・整合性ルール・数式モデル", 7)

    row = 3
    apply_section_row(ws4, row, "1. システム全体を貫く『7層のタイムアウト階層ピラミッド』（外側から内側へ）", 7)
    row += 1

    headers_s4_layer = ["レイヤー階層", "対象コンポーネント", "設定パラメータ名", "推奨設定値（例）", "役割と目的", "整合性の必須条件（大小関係）"]
    style_table_headers(ws4, row, headers_s4_layer)
    row += 1

    layers_data = [
        ("Layer 1 (最外側)", "クライアント / LB\n(ブラウザ, ALB, Nginx)", "ALB: idle_timeout\nNginx: proxy_read_timeout", "60秒", "ユーザー・外部システムとの通信切断上限。\nここが切れると画面に504 Gateway Timeoutが表示される。", "【最長であること】\nLayer 1 > Layer 2 > Layer 3"),
        ("Layer 2", "Webコンテナ\n(JBoss Undertow HTTP)", "Undertow: no-request-timeout\nrequest-timeout", "55秒", "Webリクエストスレッドの最大処理時間上限。\nLBより数秒早くタイムアウトさせ、綺麗な500/503エラーを返す。", "Layer 1 > Layer 2 > Layer 3"),
        ("Layer 3 (中核)", "JTAトランザクション\n(Narayana)", "coordinator-environment:\ndefault-timeout", "30秒", "業務トランザクション全体のマスターリミット。\n全DB・キューを巻き込んだ処理の打ち切り境界線。", "Layer 2 > Layer 3 > Layer 5\n(※Web要求時間より短く、クエリより長い)"),
        ("Layer 4", "コネクションプール取得\n(IronJacamar)", "datasource:\nblocking-timeout-millis", "5秒\n(5000ms)", "プールから空きコネクションを取得するまでの待機時間上限。\nDB高負荷時にスレッドがプール待ちで渋滞するのを防ぐ。", "Layer 3(JTA) より大幅に短く設定する\n(コネクション待ちだけでJTA時間を使い果たさない)"),
        ("Layer 5", "クエリタイムアウト\n(JDBC Statement)", "datasource: query-timeout\nset-tx-query-timeout=true", "10〜15秒\n(動的連動)", "単一SQLのDB実行上限。\nset-tx-query-timeout=true によりJTA余命に合わせて動的短縮！", "Layer 5 (Query) < Layer 3 (JTA)\nLayer 5 (Query) < Layer 7 (Socket)"),
        ("Layer 6", "XAリソースタイムアウト\n(DB側XAブランチ)", "xa-datasource:\nxa-resource-timeout", "30〜35秒", "DBエンジン自身に伝えるXAトランザクション保持上限。\nJBoss死活監視と同期させる。", "Layer 6 (XA) ≧ Layer 3 (JTA)\n※JTAより短いとDBが勝手にアボートしヒューリスティック例外！"),
        ("Layer 7 (最深部)", "ネットワークソケット\n(OS / TCP / JDBC)", "socketTimeout\noracle.net.READ_TIMEOUT", "15〜20秒\n(クエリ+3〜5秒)", "パケットの無通信検知上限。\nケーブル断線、DBハング、サイレントドロップを物理検知する。", "Layer 5 (Query) < Layer 7 (Socket)\n※ソケットが先だとSQL_CANCELパケットが送れず通信即断！")
    ]

    for item in layers_data:
        fill = fill_zebra_even if row % 2 == 0 else fill_zebra_odd
        for c_idx, val in enumerate(item, start=1):
            cell = ws4.cell(row=row, column=c_idx, value=val)
            cell.font = font_body
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = align_left_top if c_idx in [5, 6] else align_center
        ws4.row_dimensions[row].height = 56
        row += 1

    row += 2
    apply_section_row(ws4, row, "2. タイムアウト設計における『絶対不等式（ゴールデンルール）』", 6)
    row += 1

    formulas = [
        ("公式①：外側から内側への階層不等式", "LB Timeout (60s)  >  Web Request Timeout (55s)  >  JTA Tx Timeout (30s)  >  Query Timeout (15s)", "【理由】内側（DB・JTA）が先にタイムアウトして綺麗にロールバック処理を完遂した後に、上位のWebやLBがユーザーへエラー応答を返さなければならない。逆転すると『ユーザーには504エラーが出たのに、裏のDBでトランザクションが動き続けて重複更新される』という最悪のゴースト障害が発生する。"),
        ("公式②：クエリとソケットの協調不等式", "Query Timeout (15s)  <  Socket Timeout / READ_TIMEOUT (18〜20s)", "【理由】クエリタイムアウトが発火した際、JDBCドライバはDBに対して『SQL_CANCEL』という中断信号をソケット経由で送信する。もしSocket Timeoutが先に切れると、ソケット自体が物理切断されてしまい、DB側にキャンセル信号を送れず、DB側で重いSQLが動き続けてしまう。"),
        ("公式③：JTAとXAリソースの同期不等式", "xa-resource-timeout (30〜35s)  ≧  JTA default-timeout (30s)", "【理由】DB側のXAタイムアウトがJTAより短いと、Narayanaが生きていて正常処理中なのに、DBが『時間切れだ』と勝手にロールバックしてしまう。その後Narayanaがコミットを指示すると『XAER_NOTA（そんなTxは無い）』となり、データ不整合（ヒューリスティックハザード）に陥る。"),
        ("公式④：コネクション取得の早期離脱", "blocking-timeout-millis (5s)  <<  JTA default-timeout (30s)", "【理由】DBコネクションが枯渇している場合、30秒も待たせてスレッドを滞留させるのではなく、5秒程度で素早く諦めて例外を出し、後続のリクエストに枠を空ける（フェイルファスト原則）。")
    ]

    for title, formula_text, detail in formulas:
        ws4.cell(row=row, column=1, value=title).font = font_body_bold
        ws4.cell(row=row, column=1).alignment = align_center
        ws4.cell(row=row, column=1).fill = fill_sub_header
        ws4.cell(row=row, column=1).border = border_cell

        ws4.cell(row=row, column=2, value=formula_text).font = font_body_highlight
        ws4.cell(row=row, column=2).alignment = align_center
        ws4.cell(row=row, column=2).fill = fill_accent
        ws4.cell(row=row, column=2).border = border_cell

        ws4.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        for c in range(3, 7):
            cell = ws4.cell(row=row, column=c)
            cell.border = border_cell
            cell.fill = fill_zebra_odd
        top_cell = ws4.cell(row=row, column=3, value=detail)
        top_cell.font = font_body
        top_cell.alignment = align_left_top
        ws4.row_dimensions[row].height = 52
        row += 1

    adjust_column_widths(ws4, 6, {1: 24, 2: 36, 3: 20, 4: 16, 5: 38, 6: 42})

    # =========================================================================
    # SHEET 5: 設定例・CLI・トラブル対処 (Config, CLI & Troubleshooting)
    # =========================================================================
    ws5 = wb.create_sheet(title="5_JBossEAP設定例と逆引き対処")
    apply_title_row(ws5, "【実践ガイド】JBoss EAP (standalone.xml / CLI) 設定実装例 & 逆引きトラブルシューティング", 6)

    row = 3
    apply_section_row(ws5, row, "1. JBoss EAP 本番推奨 設定定義（XMLスニペット）", 6)
    row += 1

    xml_snippets = [
        ("① transactions サブシステム\n(JTAタイムアウト本体)",
         '<subsystem xmlns="urn:jboss:domain:transactions:6.0">\n'
         '    <core-environment node-identifier="${jboss.node.name}">\n'
         '        <process-id>\n'
         '            <uuid/>\n'
         '        </process-id>\n'
         '    </core-environment>\n'
         '    <recovery-environment socket-binding="txn-recovery-environment" status-socket-binding="txn-status-manager"/>\n'
         '    <coordinator-environment default-timeout="30" statistics-enabled="true"/> <!-- JTA全体タイムアウト: 30秒 -->\n'
         '    <object-store path="tx-object-store" relative-to="jboss.server.data.dir"/>\n'
         '</subsystem>',
         "【解説】\ndefault-timeout=\"30\" により、システム全体のJTAトランザクション標準タイムアウトを30秒に規定。\nstatistics-enabled=\"true\" でタイムアウト発生数(number-of-timeouts)を監視可能にする。"),
        ("② datasources サブシステム\n(XAデータソース＆クエリ連動)",
         '<xa-datasource jndi-name="java:/jdbc/MyXADS" pool-name="MyXADS_Pool" enabled="true" use-java-context="true">\n'
         '    <xa-datasource-property name="ServerName">db-server.example.com</xa-datasource-property>\n'
         '    <xa-datasource-property name="DatabaseName">orderdb</xa-datasource-property>\n'
         '    <xa-pool>\n'
         '        <min-pool-size>10</min-pool-size>\n'
         '        <max-pool-size>50</max-pool-size>\n'
         '        <is-same-rm-override>false</is-same-rm-override>\n'
         '    </xa-pool>\n'
         '    <timeout>\n'
         '        <blocking-timeout-millis>5000</blocking-timeout-millis>  <!-- プール取得待ち: 5秒 -->\n'
         '        <idle-timeout-minutes>15</idle-timeout-minutes>\n'
         '        <query-timeout>15</query-timeout>                        <!-- 単一SQLクエリ上限: 15秒 -->\n'
         '        <set-tx-query-timeout>true</set-tx-query-timeout>        <!-- JTA残り時間に動的連動！ -->\n'
         '        <xa-resource-timeout>35</xa-resource-timeout>            <!-- DB側XA制限時間: 35秒(JTA+5秒) -->\n'
         '    </timeout>\n'
         '</xa-datasource>',
         "【解説】\nset-tx-query-timeout を true にすることで、JTAトランザクション内で発行されたSQLは自動的に『JTA残り時間』がStatement.setQueryTimeoutにセットされる。\nxa-resource-timeout は JTA(30s) より少し長めの35秒にしてヒューリスティック例外を防ぐ。")
    ]

    for title, code, desc in xml_snippets:
        ws5.cell(row=row, column=1, value=title).font = font_body_bold
        ws5.cell(row=row, column=1).alignment = align_center
        ws5.cell(row=row, column=1).fill = fill_sub_header
        ws5.cell(row=row, column=1).border = border_cell

        ws5.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        for c in range(2, 5):
            cell = ws5.cell(row=row, column=c)
            cell.border = border_cell
            cell.fill = fill_zebra_odd
        code_cell = ws5.cell(row=row, column=2, value=code)
        code_cell.font = font_code
        code_cell.alignment = align_left_top

        ws5.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)
        for c in range(5, 7):
            cell = ws5.cell(row=row, column=c)
            cell.border = border_cell
            cell.fill = fill_accent
        desc_cell = ws5.cell(row=row, column=5, value=desc)
        desc_cell.font = font_body
        desc_cell.alignment = align_left_top
        ws5.row_dimensions[row].height = 135
        row += 1

    row += 2
    apply_section_row(ws5, row, "2. JBoss CLI 運用管理コマンド一覧（即時反映・設定変更）", 6)
    row += 1

    headers_s5_cli = ["操作対象", "CLIコマンド", "反映タイミング", "確認・検証コマンド"]
    # We'll use col 1 to 6
    ws5.cell(row=row, column=1, value="操作対象").font = font_header
    ws5.cell(row=row, column=1).fill = fill_header
    ws5.cell(row=row, column=1).alignment = align_center
    ws5.cell(row=row, column=1).border = border_header

    ws5.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    for c in range(2, 5):
        ws5.cell(row=row, column=c).border = border_header
        ws5.cell(row=row, column=c).fill = fill_header
    ws5.cell(row=row, column=2, value="JBoss CLI 設定変更コマンド").font = font_header
    ws5.cell(row=row, column=2).alignment = align_center

    ws5.cell(row=row, column=5, value="反映条件").font = font_header
    ws5.cell(row=row, column=5).fill = fill_header
    ws5.cell(row=row, column=5).alignment = align_center
    ws5.cell(row=row, column=5).border = border_header

    ws5.cell(row=row, column=6, value="確認・メトリクス取得コマンド").font = font_header
    ws5.cell(row=row, column=6).fill = fill_header
    ws5.cell(row=row, column=6).alignment = align_center
    ws5.cell(row=row, column=6).border = border_header
    ws5.row_dimensions[row].height = 28
    row += 1

    cli_data = [
        ("JTAタイムアウト変更",
         "/subsystem=transactions:write-attribute(name=default-timeout, value=30)",
         "即時反映\n(新規開始Txから適用)",
         "/subsystem=transactions:read-attribute(name=default-timeout)\n/subsystem=transactions:read-attribute(name=number-of-timeouts)"),
        ("set-tx-query-timeout 有効化",
         "/subsystem=datasources/xa-data-source=MyXADS_Pool:write-attribute(name=set-tx-query-timeout, value=true)",
         "データソース再読込\nまたはreload",
         "/subsystem=datasources/xa-data-source=MyXADS_Pool:read-attribute(name=set-tx-query-timeout)"),
        ("query-timeout 変更",
         "/subsystem=datasources/xa-data-source=MyXADS_Pool:write-attribute(name=query-timeout, value=15)",
         "即時反映",
         "/subsystem=datasources/xa-data-source=MyXADS_Pool:read-attribute(name=query-timeout)"),
        ("xa-resource-timeout 変更",
         "/subsystem=datasources/xa-data-source=MyXADS_Pool:write-attribute(name=xa-resource-timeout, value=35)",
         "データソース再読込\nまたはreload",
         "/subsystem=datasources/xa-data-source=MyXADS_Pool:read-attribute(name=xa-resource-timeout)")
    ]

    for target, cmd, timing, verify in cli_data:
        fill = fill_zebra_even if row % 2 == 0 else fill_zebra_odd
        ws5.cell(row=row, column=1, value=target).font = font_body_bold
        ws5.cell(row=row, column=1).alignment = align_center
        ws5.cell(row=row, column=1).fill = fill
        ws5.cell(row=row, column=1).border = border_cell

        ws5.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        for c in range(2, 5):
            ws5.cell(row=row, column=c).border = border_cell
            ws5.cell(row=row, column=c).fill = fill
        c_cell = ws5.cell(row=row, column=2, value=cmd)
        c_cell.font = font_code
        c_cell.alignment = align_left_top

        ws5.cell(row=row, column=5, value=timing).font = font_body
        ws5.cell(row=row, column=5).alignment = align_center
        ws5.cell(row=row, column=5).fill = fill
        ws5.cell(row=row, column=5).border = border_cell

        v_cell = ws5.cell(row=row, column=6, value=verify)
        v_cell.font = font_code
        v_cell.alignment = align_left_top
        v_cell.fill = fill
        v_cell.border = border_cell
        ws5.row_dimensions[row].height = 42
        row += 1

    row += 2
    apply_section_row(ws5, row, "3. 逆引きトラブルシューティング（エラーログ・根本原因・解決策）", 6)
    row += 1

    headers_s5_trouble = ["ログ・エラーメッセージ", "発生フェーズ・条件", "根本原因 (Root Cause)", "即時対処と恒久対策"]
    # We will use cols 1 to 6
    ws5.cell(row=row, column=1, value="代表的エラーログ").font = font_header
    ws5.cell(row=row, column=1).fill = fill_header
    ws5.cell(row=row, column=1).alignment = align_center
    ws5.cell(row=row, column=1).border = border_header

    ws5.cell(row=row, column=2, value="発生フェーズ").font = font_header
    ws5.cell(row=row, column=2).fill = fill_header
    ws5.cell(row=row, column=2).alignment = align_center
    ws5.cell(row=row, column=2).border = border_header

    ws5.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
    for c in range(3, 5):
        ws5.cell(row=row, column=c).border = border_header
        ws5.cell(row=row, column=c).fill = fill_header
    ws5.cell(row=row, column=3, value="根本原因 (Root Cause)").font = font_header
    ws5.cell(row=row, column=3).alignment = align_center

    ws5.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)
    for c in range(5, 7):
        ws5.cell(row=row, column=c).border = border_header
        ws5.cell(row=row, column=c).fill = fill_header
    ws5.cell(row=row, column=5, value="即時対処と恒久対策").font = font_header
    ws5.cell(row=row, column=5).alignment = align_center
    ws5.row_dimensions[row].height = 28
    row += 1

    trouble_data = [
        (
            "ARJUNA016029: Transaction has timed out; action was cancelled\n"
            "ARJUNA012117: TransactionReaper::check timeout",
            "Phase 0\n(業務ロジック実行中)",
            "トランザクション開始後の経過時間が default-timeout を超過したため、NarayanaのTransactionReaperが強制ロールバックを発行した。",
            "①遅いSQLや外部API通信がないかAPM/ログで特定。\n②set-tx-query-timeout=true を設定してSQLが余命を超えて暴走するのを抑止。\n③業務的に長時間必要な場合は該当EJB/メソッドのみ @TransactionTimeout で個別拡大。"
        ),
        (
            "ARJUNA012378: ReaperElement appears to be wedged: ActionStatus.RUNNING\n"
            "ARJUNA012121: TransactionReaper::doCancellations worker Zombie",
            "Phase 0 / Phase 1\n(Reaper打切り時)",
            "TransactionReaperがTxをロールバックしようとしたが、スレッドがDBソケット待ち(socketRead0)等で停止しており、cancel-fail-wait-period 内に中断できなかった。",
            "①【最重要】JDBCドライバの socketTimeout / READ_TIMEOUT を設定し、ソケットが永遠にブロックするのを防ぐ！\n②cancel-wait-period / cancel-fail-wait-period をデフォルト(500ms)から1000〜2000msへ調整検討。"
        ),
        (
            "javax.transaction.xa.XAException: XAER_NOTA\n"
            "(XID not found)",
            "Phase 1 または Phase 2\n(Prepare / Commit時)",
            "DB側のタイムアウト(xa-resource-timeout等)がJTAタイムアウトより先に切れて、DBが自律的にトランザクションブランチを消去してしまった。",
            "①【不等式の確認】xa-resource-timeout ≧ JTA default-timeout になっているか確認！\n②DB側のアイドルタイムアウト設定やOracleの分散トランザクション設定を見直し、JTAより先に消えないよう調整。"
        ),
        (
            "javax.transaction.HeuristicMixedException\n"
            "または HeuristicRollbackException",
            "Phase 2\n(Commit完了フェーズ)",
            "2PCの最悪障害。複数DBのうち、DB-Aはコミット成功したが、DB-Bはタイムアウト等でロールバックされてしまい、データが不整合状態になった。",
            "①DB側の未解決XAトランザクションテーブル (DBA_2PC_PENDING等) をDBAと確認し手動リカバリ。\n②JBossのRecoveryManagerが正常稼働しているか確認。\n③タイムアウト値の整合性（JTA < XA-Resource < Socket）を厳格に再設計。"
        ),
        (
            "ORA-25351: timeout waiting for transaction to be suspended\n"
            "(Oracle XA Resume Timeout)",
            "コネクション再利用\n(XA RESUME時)",
            "スレッドが別スレッドの保留したXAブランチを resume(TMRESUME) しようとしたが、相手側セッションが suspend しないまま指定時間が経過した。",
            "①同一トランザクションを複数スレッドで同時並行利用していないかアプリ設計を見直し。\n②IronJacamarのインターリーブ設定やプール共有設定を確認。\n③DBMS_TRANSACTION.RESUME の待機時間を考慮。"
        )
    ]

    for err, phase, cause, fix in trouble_data:
        fill = fill_zebra_even if row % 2 == 0 else fill_zebra_odd
        ws5.cell(row=row, column=1, value=err).font = font_code
        ws5.cell(row=row, column=1).alignment = align_left_top
        ws5.cell(row=row, column=1).fill = fill
        ws5.cell(row=row, column=1).border = border_cell

        ws5.cell(row=row, column=2, value=phase).font = font_body_bold
        ws5.cell(row=row, column=2).alignment = align_center
        ws5.cell(row=row, column=2).fill = fill
        ws5.cell(row=row, column=2).border = border_cell

        ws5.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        for c in range(3, 5):
            ws5.cell(row=row, column=c).border = border_cell
            ws5.cell(row=row, column=c).fill = fill
        c_cell = ws5.cell(row=row, column=3, value=cause)
        c_cell.font = font_body
        c_cell.alignment = align_left_top

        ws5.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)
        for c in range(5, 7):
            ws5.cell(row=row, column=c).border = border_cell
            ws5.cell(row=row, column=c).fill = fill
        f_cell = ws5.cell(row=row, column=5, value=fix)
        f_cell.font = font_body
        f_cell.alignment = align_left_top
        ws5.row_dimensions[row].height = 68
        row += 1

    adjust_column_widths(ws5, 6, {1: 34, 2: 18, 3: 20, 4: 20, 5: 32, 6: 34})

    # Freeze panes
    ws1.freeze_panes = "A2"
    ws2.freeze_panes = "A5"
    ws3.freeze_panes = "A2"
    ws4.freeze_panes = "A5"
    ws5.freeze_panes = "A2"

    # Save workbook
    output_filename = "JTA_Transaction_Timeout_and_2PC_Master_Guide.xlsx"
    wb.save(output_filename)
    print(f"Successfully generated: {output_filename}")

if __name__ == "__main__":
    build_workbook()
