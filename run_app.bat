@echo off
REM ���̃X�N���v�g�́APython�̊��ϐ����ݒ肳��Ă���Windows���ŁA
REM 'tag_classification_app_tkinter.py' �A�v���P�[�V���������s���܂��B

REM �K�v�ȃ��C�u���� (pandas) ���C���X�g�[�����܂��B
REM ���łɃC���X�g�[������Ă���ꍇ�̓X�L�b�v����܂��B
echo �K�v�ȃ��C�u�������C���X�g�[�����Ă��܂�...
pip install pandas
if %errorlevel% neq 0 (
    echo.
    echo �G���[: pandas�̃C���X�g�[���Ɏ��s���܂����B
    echo Python���������C���X�g�[������A���ϐ�PATH���ݒ肳��Ă��邩�m�F���Ă��������B
    echo.
    pause
    exit /b %errorlevel%
)
echo �C���X�g�[�������B

REM Python�X�N���v�g�����s
echo �A�v���P�[�V�������N�����Ă��܂�...
python tag_classification_app_tkinter.py

REM �A�v���P�[�V����������ꂽ��A���̃E�B���h�E�����̂�҂�
pause
