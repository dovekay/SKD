const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  LevelFormat, PageNumber, Footer, BorderStyle
} = require('docx');
const fs = require('fs');

const today = '2025-04-25';

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, font: 'Arial' })]
  });
}

function p(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: 'Arial' })]
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
        children: [new TextRun({ text: '서비스 이용약관', bold: true, size: 48, font: 'Arial', color: '1F3864' })]
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

      h2('제1조 (목적)'),
      p('본 약관은 dovekay(이하 "운영자")가 제공하는 KE Schedule Uploader(이하 "서비스")의 이용 조건 및 절차, 운영자와 이용자의 권리·의무 및 책임 사항을 규정함을 목적으로 합니다.'),

      h2('제2조 (약관의 효력 및 변경)'),
      p('① 본 약관은 서비스를 이용하고자 하는 모든 이용자에게 적용됩니다.'),
      p('② 운영자는 관련 법령을 위배하지 않는 범위에서 본 약관을 개정할 수 있으며, 변경된 약관은 서비스 내 공지 또는 앱 스토어 업데이트 노트를 통해 사전 고지합니다.'),
      p('③ 이용자가 변경된 약관에 동의하지 않을 경우 서비스 이용을 중단하고 연동을 해제할 수 있습니다.'),

      h2('제3조 (서비스 내용)'),
      p('서비스는 사용자가 KE(대한항공) 스케줄 정보를 Google 캘린더에 자동으로 업로드할 수 있도록 지원합니다. 서비스는 Google Calendar API를 통해 이용자의 캘린더에 일정을 읽고 쓸 수 있습니다.'),

      h2('제4조 (서비스 이용)'),
      p('① 서비스를 이용하려면 유효한 Google 계정이 필요하며, 이용자는 Google OAuth 2.0을 통해 서비스에 캘린더 접근 권한을 부여해야 합니다.'),
      p('② 이용자는 다음 행위를 해서는 안 됩니다:'),
      bullet('서비스의 정상적인 운영을 방해하는 행위'),
      bullet('타인의 개인정보를 무단으로 수집·이용하는 행위'),
      bullet('서비스를 이용하여 법령 또는 공공질서에 위반되는 행위'),
      bullet('서비스를 역공학(리버스 엔지니어링)하거나 무단으로 복제하는 행위'),

      h2('제5조 (서비스의 중단)'),
      p('운영자는 다음과 같은 경우 서비스 제공을 일시적으로 중단할 수 있습니다:'),
      bullet('시스템 점검, 교체, 고장 등 기술적 사유'),
      bullet('Google API 정책 변경 또는 서비스 중단'),
      bullet('천재지변, 국가비상사태 등 불가항력적 사유'),
      blank(),
      p('서비스 중단 시 가능한 경우 사전 공지하며, 불가피한 경우 사후 공지합니다.'),

      h2('제6조 (면책 조항)'),
      p('① 운영자는 무료로 제공되는 서비스의 이용과 관련하여 관련 법령에 특별한 규정이 없는 한 책임을 지지 않습니다.'),
      p('② 운영자는 이용자의 귀책사유로 인한 서비스 이용 장애에 대해 책임을 지지 않습니다.'),
      p('③ 운영자는 Google Calendar API 정책 변경으로 인해 서비스 기능이 제한되는 경우에 대해 책임을 지지 않습니다.'),
      p('④ 서비스는 "있는 그대로(AS-IS)" 제공되며, 특정 목적에의 적합성에 대한 보증을 하지 않습니다.'),

      h2('제7조 (준거법 및 분쟁 해결)'),
      p('본 약관은 대한민국 법률에 따라 해석되며, 서비스 이용과 관련하여 분쟁이 발생할 경우 운영자와 이용자가 성실하게 협의하여 해결합니다. 협의가 이루어지지 않을 경우 민사소송법상 관할 법원에 소를 제기할 수 있습니다.'),

      h2('제8조 (연락처)'),
      p('운영자: dovekay'),
      p('서비스 내 문의 기능 또는 앱 스토어 개발자 연락처를 통해 문의하시기 바랍니다.'),

      divider(),

      // ─── ENGLISH ───
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 160 },
        children: [new TextRun({ text: 'Terms of Service', bold: true, size: 48, font: 'Arial', color: '1F3864' })]
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

      h2('Article 1 – Purpose'),
      p('These Terms of Service govern the conditions and procedures for using KE Schedule Uploader (the "Service") operated by dovekay (the "Operator"), and set forth the rights, obligations, and responsibilities of the Operator and users.'),

      h2('Article 2 – Effect and Amendment of Terms'),
      p('① These Terms apply to all users who wish to use the Service.'),
      p('② The Operator may amend these Terms to the extent permitted by applicable law. Any changes will be announced in advance through in-app notices or app store release notes.'),
      p('③ If a user does not agree to the amended Terms, the user may stop using the Service and revoke the account connection.'),

      h2('Article 3 – Description of Service'),
      p('The Service helps users automatically upload KE (Korean Air) schedule information to Google Calendar. It uses the Google Calendar API to read and write calendar events on behalf of the user.'),

      h2('Article 4 – Use of Service'),
      p('① A valid Google account is required to use the Service. Users must grant calendar access permissions to the Service via Google OAuth 2.0.'),
      p('② Users must not:'),
      bullet('Interfere with or disrupt the normal operation of the Service'),
      bullet('Collect or use other persons\' personal information without authorization'),
      bullet('Use the Service for any purpose that violates applicable law or public order'),
      bullet('Reverse engineer, decompile, or reproduce the Service without authorization'),

      h2('Article 5 – Service Interruption'),
      p('The Operator may temporarily suspend the Service in the following circumstances:'),
      bullet('Technical reasons such as system maintenance, replacement, or failure'),
      bullet('Changes to or suspension of the Google API'),
      bullet('Force majeure events such as natural disasters or national emergencies'),
      blank(),
      p('The Operator will provide advance notice of interruptions where possible, and post-hoc notice when advance notice is not feasible.'),

      h2('Article 6 – Disclaimer'),
      p('① Except as required by applicable law, the Operator is not liable for issues arising from use of the free Service.'),
      p('② The Operator is not liable for service disruptions caused by the user\'s own actions.'),
      p('③ The Operator is not liable for limitations in Service functionality caused by changes to Google Calendar API policies.'),
      p('④ Service is provided "AS IS" without any warranty of fitness for a particular purpose.'),

      h2('Article 7 – Governing Law and Dispute Resolution'),
      p('These Terms are governed by the laws of the Republic of Korea. In the event of a dispute, the Operator and user shall seek resolution through good-faith negotiation. If negotiation fails, the dispute may be submitted to the competent court under the Civil Procedure Act.'),

      h2('Article 8 – Contact'),
      p('Operator: dovekay'),
      p('Please contact us using the in-app contact feature or the developer contact on the app store listing.'),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('KE_Schedule_Uploader_TermsOfService.docx', buf);
  console.log('Done: Terms of Service');
});