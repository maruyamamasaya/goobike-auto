import Head from "next/head";
import Link from "next/link";
import type { GetStaticProps, InferGetStaticPropsType } from "next";
import fs from "fs";
import path from "path";

const SCRIPT_PATH = path.join(process.cwd(), "public", "scripts", "goobikeAutoFill.js");

type Props = {
  script: string;
};

export const getStaticProps: GetStaticProps<Props> = async () => {
  const script = await fs.promises.readFile(SCRIPT_PATH, "utf8");
  return {
    props: { script },
  };
};

export default function AutoFillPage({
  script,
}: InferGetStaticPropsType<typeof getStaticProps>) {
  return (
    <>
      <Head>
        <title>Goobike Auto - スプレッドシート取り込みスクリプト</title>
        <meta
          name="description"
          content="グーバイク登録フォームへスプレッドシートの値を貼り付ける自動入力スクリプト"
        />
      </Head>
      <main style={{ padding: "2rem", fontFamily: "system-ui, sans-serif", maxWidth: "960px" }}>
        <Link href="/">← ダッシュボードに戻る</Link>
        <h1 style={{ marginTop: "1rem" }}>スプレッドシート → フォーム自動入力</h1>
        <p style={{ lineHeight: 1.8 }}>
          以下のスクリプトを Chrome DevTools などのコンソールに貼り付けると、
          スプレッドシートからコピーしたヘッダー行とデータ行を自動的に解釈し、
          GooBike の登録フォームに値を流し込みます。
          事前に GooBike の対象ページを開きログインしておいてください。
        </p>

        <section style={{ marginTop: "2rem" }}>
          <h2>使い方</h2>
          <ol style={{ lineHeight: 1.8 }}>
            <li>スプレッドシートでヘッダー行と登録したい行をコピーします（2行）。</li>
            <li>GooBike の登録フォームを開き、ブラウザのコンソールを開きます。</li>
            <li>下記スクリプトを貼り付けて Enter を押します。</li>
            <li>表示されるダイアログにコピーした内容を貼り付けると、自動で入力されます。</li>
          </ol>
        </section>

        <section style={{ marginTop: "2rem" }}>
          <h2>スクリプト</h2>
          <div
            style={{
              marginTop: "1rem",
              padding: "1rem",
              borderRadius: "8px",
              background: "#111",
              color: "#f9fafb",
              overflowX: "auto",
            }}
          >
            <pre style={{ margin: 0, fontSize: "0.85rem", lineHeight: 1.5 }}>
              <code>{script}</code>
            </pre>
          </div>
        </section>

        <section style={{ marginTop: "2rem" }}>
          <h2>主な機能</h2>
          <ul style={{ lineHeight: 1.8 }}>
            <li>ヘッダー名を元にマッピングするため、列の入れ替えに強い設計です。</li>
            <li>令和・平成・昭和などの和暦表記を西暦に自動変換します。</li>
            <li>走行距離・価格・オプションなどの入力ルールを内蔵しています。</li>
            <li>未入力や「選択してください」などの選択状態は自動スキップします。</li>
          </ul>
        </section>
      </main>
    </>
  );
}
