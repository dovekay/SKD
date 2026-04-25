const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  LevelFormat, PageNumber, Footer, NumberFormat, BorderStyle
} = require('docx');
const fs = require('fs');

const today = '2025-04-25';

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, font: 'Arial' })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, font: 'Arial' })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: 'Arial', ...opts })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: 'Arial' })]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E75B6', space: 1 } },
    children: []
  });
}

function blank() {
  return new Paragraph({ children: [new TextRun('')] });
}

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 22 } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: '1F3864' },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial', color: '2E75B6' },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: 'Arial', color: '888888' })
          ]
        })]
      })
    },
    children: [
      // ─── KOREAN ───
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 160 },
        children: [new TextRun({ text: '개인정보처리방침', bold: true, size: 48, font: 'Arial', color: '1F3864' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: 'KE Schedule Uploader', size: 26, font: 'Arial', color: '2E75B6' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: `시행일: ${today}`, size: 22, font: 'Arial', color: '666666' })]
      }),

      h2('1. 개요'),
      p('KE Schedule Uploader(이하 "서비스")는 dovekay(이하 "운영자")가 제공하는 서비스입니다. 본 개인정보처리방침은 서비스 이용 과정에서 수집·이용·보관·파기되는 개인정보에 관한 사항을 규정합니다.'),

      h2('2. 수집하는 개인정보 항목 및 수집 방법'),
      p('서비스는 구글 OAuth 2.0 인증을 통해 다음 정보에 접근합니다:'),
      bullet('Google 계정 기본 프로필 정보 (이름, 이메일 주소)'),
      bullet('Google 캘린더 일정 데이터 (읽기 및 쓰기 권한)'),
      blank(),
      p('위 정보는 사용자가 구글 계정 연동을 직접 승인하는 시점에 수집됩니다.'),

      h2('3. 개인정보의 이용 목적'),
      p('수집된 정보는 다음 목적에만 사용됩니다:'),
      bullet('Google 캘린더에 일정을 업로드하는 핵심 서비스 기능 제공'),
      bullet('사용자 인증 및 세션 유지'),
      blank(),
      p('수집된 정보는 제3자에게 제공·판매·공유되지 않으며, 광고 목적으로 사용되지 않습니다.'),

      h2('4. 개인정보의 보유 및 이용 기간'),
      p('서비스는 최소한의 정보만 처리하며, 별도의 서버에 개인정보를 저장하지 않습니다. OAuth 액세스 토큰은 서비스 기능 수행을 위해 세션 동안만 사용되며, 사용자가 Google 계정 연동을 해제하면 즉시 무효화됩니다.'),

      h2('5. 개인정보의 제3자 제공'),
      p('운영자는 원칙적으로 이용자의 개인정보를 외부에 제공하지 않습니다. 단, 다음의 경우에는 예외로 합니다:'),
      bullet('이용자가 사전에 동의한 경우'),
      bullet('법령의 규정에 의거하거나 수사 목적으로 법령에 정해진 절차와 방법에 따라 수사기관의 요구가 있는 경우'),

      h2('6. 쿠키(Cookie) 및 유사 기술의 사용'),
      p('서비스는 세션 유지 및 인증 상태 관리를 위해 브라우저 로컬 스토리지 또는 세션 스토리지를 사용할 수 있습니다. 사용자는 브라우저 설정을 통해 이를 거부할 수 있으나, 일부 서비스 기능이 제한될 수 있습니다.'),

      h2('7. 이용자의 권리 및 행사 방법'),
      p('이용자는 언제든지 다음 권리를 행사할 수 있습니다:'),
      bullet('Google 계정 설정(myaccount.google.com)에서 본 서비스의 접근 권한 철회'),
      bullet('개인정보 열람, 정정, 삭제 요청'),
      blank(),
      p('권리 행사는 아래 연락처로 요청하실 수 있습니다.'),

      h2('8. 개인정보 보호책임자'),
      p('운영자: dovekay'),
      p('서비스 내 문의 기능 또는 구글 플레이 개발자 연락처를 통해 문의하시기 바랍니다.'),

      h2('9. 방침 변경'),
      p('본 개인정보처리방침은 관련 법령이나 서비스 변경에 따라 개정될 수 있습니다. 변경 시에는 서비스 내 공지 또는 앱 스토어 업데이트 노트를 통해 사전 고지합니다.'),

      divider(),

      // ─── ENGLISH ───
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 160 },
        children: [new TextRun({ text: 'Privacy Policy', bold: true, size: 48, font: 'Arial', color: '1F3864' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: 'KE Schedule Uploader', size: 26, font: 'Arial', color: '2E75B6' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: `Effective Date: ${today}`, size: 22, font: 'Arial', color: '666666' })]
      }),

      h2('1. Overview'),
      p('KE Schedule Uploader (the "Service") is operated by dovekay (the "Operator"). This Privacy Policy describes how we collect, use, store, and delete personal information when you use the Service.'),

      h2('2. Information We Collect'),
      p('The Service accesses the following information through Google OAuth 2.0 authentication:'),
      bullet('Basic Google account profile (name and email address)'),
      bullet('Google Calendar event data (read and write access)'),
      blank(),
      p('This information is collected at the time the user explicitly authorizes the Google account connection.'),

      h2('3. How We Use Your Information'),
      p('Collected information is used solely for the following purposes:'),
      bullet('Providing the core service functionality of uploading events to Google Calendar'),
      bullet('User authentication and session management'),
      blank(),
      p('We do not sell, share, or otherwise disclose collected information to third parties, nor do we use it for advertising purposes.'),

      h2('4. Data Retention'),
      p('The Service processes only the minimum necessary information and does not store personal data on any external server. OAuth access tokens are used only during the active session to perform service functions and are immediately invalidated when the user revokes Google account access.'),

      h2('5. Sharing of Personal Information'),
      p('The Operator does not share user personal information with external parties, except in the following circumstances:'),
      bullet('When the user has given prior consent'),
      bullet('When required by law or requested by law enforcement authorities in accordance with applicable legal procedures'),

      h2('6. Cookies and Similar Technologies'),
      p("The Service may use browser local storage or session storage to maintain session state and authentication. Users may opt out through browser settings; however, certain Service features may become unavailable."),

      h2('7. Your Rights'),
      p('You may exercise the following rights at any time:'),
      bullet("Revoke the Service's access permissions via Google Account settings (myaccount.google.com)"),
      bullet('Request access to, correction of, or deletion of your personal information'),
      blank(),
      p('To exercise these rights, please contact us using the information below.'),

      h2('8. Data Controller'),
      p('Operator: dovekay'),
      p('Please use the in-app contact feature or the developer contact on the app store listing.'),

      h2('9. Changes to This Policy'),
      p('This Privacy Policy may be updated to reflect changes in applicable laws or the Service. We will notify users of any material changes via in-app notice or app store release notes prior to the changes taking effect.'),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('KE_Schedule_Uploader_PrivacyPolicy.docx', buf);
  console.log('Done: Privacy Policy');
});