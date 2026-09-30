# 소개페이지 (Vercel 배포용)

정적 HTML 단일 페이지입니다. 이 폴더만 배포하면 됩니다.

## 배포 방법 (CLI)
```
cd intro-vercel
npx vercel        # 미리보기 배포
npx vercel --prod # 운영 배포
```
프로젝트 설정 질문에서 Framework는 `Other`, Build Command/Output Directory는 비워두세요.

## 배포 방법 (GitHub 연동)
Vercel 대시보드 > Add New Project > 저장소 선택 > **Root Directory를 `intro-vercel`로 지정**.
