@echo off
  
rem
rem bin �t�H���_�� xlsm �t�@�C������ VBA �X�N���v�g���ꊇ�G�N�X�|�[�g���܂��B
rem �����t�H���_�ɁAvbac.wsf�Abin �t�H���_���K�v�ł��B
rem xlsm �t�@�C���́Abin �t�H���_�Ɋi�[���܂��B
rem
  
rem ���̃o�b�`�����݂���t�H���_���J�����g�ɐݒ�
pushd %0\..
cls
  
rem �G�N�X�|�[�g
cscript vbac.wsf decombine
  
exit