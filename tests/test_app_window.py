import pytest
import configparser
from datetime import datetime
from unittest.mock import MagicMock, patch

from app_window import MedicalDocsAnalyzerGUI
from config_manager import load_config


def restore_config(config, original_config):
    """configを元の状態に復元するヘルパーメソッド"""
    for section in config.sections():
        config.remove_section(section)
    for section in original_config.sections():
        if not config.has_section(section):
            config.add_section(section)
        for key, value in original_config[section].items():
            config[section][key] = value


@pytest.fixture
def original_config():
    """オリジナルのconfigを保存するフィクスチャ"""
    return load_config()


@pytest.fixture
def mock_tk():
    """tkinterのモックを作成するフィクスチャ"""
    with patch('app_window.tk') as mock_tk:
        # Tkオブジェクトのモック
        mock_root = MagicMock()
        mock_tk.Tk.return_value = mock_root

        # ttk関連のモック
        mock_tk.ttk = MagicMock()
        mock_frame = MagicMock()
        mock_label_frame = MagicMock()
        mock_label = MagicMock()
        mock_button = MagicMock()

        mock_tk.ttk.Frame.return_value = mock_frame
        mock_tk.ttk.LabelFrame.return_value = mock_label_frame
        mock_tk.ttk.Label.return_value = mock_label
        mock_tk.ttk.Button.return_value = mock_button

        # tkの定数
        mock_tk.W = "w"
        mock_tk.E = "e"
        mock_tk.N = "n"
        mock_tk.S = "s"

        yield mock_tk


@pytest.fixture
def mock_date_entry():
    """DateEntryのモックを作成するフィクスチャ"""
    with patch('app_window.DateEntry') as mock_date_entry:
        mock_date = MagicMock()
        mock_date.get_date.return_value = datetime.now().date()
        mock_date_entry.return_value = mock_date
        yield mock_date_entry


@pytest.fixture
def gui(original_config):
    """GUIインスタンスを作成するフィクスチャ"""
    # すべての外部依存をモック化
    with patch('app_window.tk'), \
            patch('app_window.ttk'), \
            patch('app_window.DateEntry'), \
            patch('app_window.process_medical_documents'), \
            patch('app_window.messagebox'), \
            patch('app_window.MedicalDocsAnalyzer') as mock_analyzer:
        # tkとttkのモックを取得
        mock_tk = pytest.importorskip('app_window').tk
        mock_ttk = pytest.importorskip('app_window').ttk

        # ルートとフレームのモック
        mock_root = MagicMock()
        mock_frame = MagicMock()
        mock_label_frame = MagicMock()

        # tkの定数を設定
        mock_tk.W = "w"
        mock_tk.E = "e"
        mock_tk.N = "n"
        mock_tk.S = "s"

        # モックを設定
        mock_tk.Tk.return_value = mock_root
        mock_ttk.Frame.return_value = mock_frame
        mock_ttk.LabelFrame.return_value = mock_label_frame

        # DateEntryのモック
        mock_date = MagicMock()
        mock_date.get_date.return_value = datetime.now().date()
        pytest.importorskip('app_window').DateEntry.return_value = mock_date

        # GUIのインスタンスを作成
        gui = MedicalDocsAnalyzerGUI(mock_root)

        # _setup_buttonsメソッドが実際に呼び出されたかを確認するためのスパイを追加
        gui._setup_buttons_called = mock_ttk.Button.call_count

        yield gui

        # テスト後に設定を元に戻す
        restore_config(gui.config, original_config)


class TestMedicalDocsAnalyzerGUI:

    def test_init(self, gui):
        """初期化が正しく行われるかテスト"""
        assert gui.root is not None
        # titleがセットされたことを確認
        gui.root.title.assert_called_once()
        assert gui.config is not None
        assert gui.analyzer is not None

    def test_setup_date_frame(self, gui):
        """日付フレームの設定が正しく行われるかテスト"""
        assert hasattr(gui, 'start_date')
        assert hasattr(gui, 'end_date')

        # DateEntryが呼ばれたことを確認
        assert pytest.importorskip('app_window').DateEntry.call_count >= 2

    def test_setup_buttons(self, gui):
        """ボタンの設定が正しく行われるかテスト"""
        # _setup_buttonsメソッドが呼ばれたことを確認する代替方法
        buttons_called = gui._setup_buttons_called > 0
        assert buttons_called, "ボタンの設定メソッドが呼ばれていない"

    @patch('app_window.process_medical_documents')
    @patch('app_window.messagebox')
    def test_load_data_success(self, mock_messagebox, mock_process, gui):
        """データ読込の成功パターンをテスト"""
        mock_process.return_value = True

        # save_date_to_configをモックに置き換え
        with patch.object(gui, 'save_date_to_config', return_value=True):
            gui.load_data()

            # process_medical_documentsが呼ばれたことを確認
            mock_process.assert_called_once()

            # 成功メッセージが表示されたことを確認
            mock_messagebox.showinfo.assert_called_once_with("完了", "データ読込が完了しました。")

    @patch('app_window.process_medical_documents')
    @patch('app_window.messagebox')
    def test_load_data_failure(self, mock_messagebox, mock_process, gui):
        """データ読込の失敗パターンをテスト"""
        mock_process.return_value = False

        # save_date_to_configをモックに置き換え
        with patch.object(gui, 'save_date_to_config', return_value=True):
            gui.load_data()

            # process_medical_documentsが呼ばれたことを確認
            mock_process.assert_called_once()

            # エラーメッセージが表示されたことを確認
            mock_messagebox.showerror.assert_called_once_with("エラー", "データ読込中にエラーが発生しました。")

    @patch('app_window.process_medical_documents')
    @patch('app_window.messagebox')
    def test_load_data_exception(self, mock_messagebox, mock_process, gui):
        """データ読込でExceptionが発生した場合のテスト"""
        mock_process.side_effect = Exception("テストエラー")

        # save_date_to_configをモックに置き換え
        with patch.object(gui, 'save_date_to_config', return_value=True):
            gui.load_data()

            # エラーメッセージが表示されたことを確認
            mock_messagebox.showerror.assert_called_once()
            assert "予期せぬエラー" in mock_messagebox.showerror.call_args[0][1]

    @patch('app_window.save_config')
    @patch('app_window.messagebox')
    def test_save_date_to_config_success(self, mock_messagebox, mock_save_config, gui):
        """設定の保存が成功する場合のテスト"""
        # 開始日と終了日のget_dateが同じ日付を返すようにモック化済み
        result = gui.save_date_to_config()

        assert result is True
        mock_save_config.assert_called_once_with(gui.config)
        assert 'Analysis' in gui.config
        assert 'start_date' in gui.config['Analysis']
        assert 'end_date' in gui.config['Analysis']

    @patch('app_window.save_config')
    @patch('app_window.messagebox')
    def test_save_date_to_config_invalid_date(self, mock_messagebox, mock_save_config, gui):
        """終了日より後の開始日を設定した場合のテスト"""
        # 開始日が終了日より後になるようにモック
        with patch.object(gui, 'start_date') as mock_start_date, \
                patch.object(gui, 'end_date') as mock_end_date:
            mock_start_date.get_date.return_value = datetime(2025, 2, 1).date()
            mock_end_date.get_date.return_value = datetime(2025, 1, 1).date()

            result = gui.save_date_to_config()

            assert result is False
            mock_messagebox.showerror.assert_called_once()
            assert "開始日が終了日より後の日付" in mock_messagebox.showerror.call_args[0][1]
            mock_save_config.assert_not_called()

    @patch('app_window.messagebox')
    def test_start_analysis_success(self, mock_messagebox, gui):
        """分析開始が成功する場合のテスト"""
        # save_date_to_configをモックに置き換え
        with patch.object(gui, 'save_date_to_config', return_value=True):
            # analyzerのrun_analysisをモックに置き換え
            gui.analyzer.run_analysis = MagicMock(return_value=(True, "成功"))

            gui.start_analysis()

            # run_analysisが呼ばれたことを確認
            gui.analyzer.run_analysis.assert_called_once()

            # エラーメッセージが表示されなかったことを確認
            mock_messagebox.showerror.assert_not_called()

    @patch('app_window.messagebox')
    def test_start_analysis_date_error(self, mock_messagebox, gui):
        """日付設定でエラーが発生した場合のテスト"""
        # save_date_to_configをモックに置き換え
        with patch.object(gui, 'save_date_to_config', return_value=False):
            # analyzerのrun_analysisをモックに置き換え
            gui.analyzer.run_analysis = MagicMock()

            gui.start_analysis()

            # run_analysisが呼ばれなかったことを確認
            gui.analyzer.run_analysis.assert_not_called()

    @patch('app_window.messagebox')
    def test_start_analysis_exception(self, mock_messagebox, gui):
        """分析実行中に例外が発生した場合のテスト"""
        # save_date_to_configをモックに置き換え
        with patch.object(gui, 'save_date_to_config', return_value=True):
            # analyzerのrun_analysisをモックに置き換え
            gui.analyzer.run_analysis = MagicMock(side_effect=Exception("テストエラー"))

            gui.start_analysis()

            # run_analysisが呼ばれたことを確認
            gui.analyzer.run_analysis.assert_called_once()

            # エラーメッセージが表示されたことを確認
            mock_messagebox.showerror.assert_called_once()
            assert "予期せぬエラー" in mock_messagebox.showerror.call_args[0][1]

    @patch('app_window.subprocess.Popen')
    def test_open_config_success(self, mock_popen, gui):
        """設定ファイルを開くテスト"""
        gui.open_config()

        # subprocessが呼ばれたことを確認
        mock_popen.assert_called_once()
        args = mock_popen.call_args[0][0]
        assert args[0] == 'notepad.exe'
        assert 'config_path' in gui.config['PATHS']

    @patch('app_window.subprocess.Popen')
    @patch('app_window.messagebox')
    def test_open_config_error(self, mock_messagebox, mock_popen, gui):
        """設定ファイルを開くエラーのテスト"""
        mock_popen.side_effect = Exception("テストエラー")

        gui.open_config()

        # エラーメッセージが表示されたことを確認
        mock_messagebox.showerror.assert_called_once()
        assert "設定ファイルを開けませんでした" in mock_messagebox.showerror.call_args[0][1]
