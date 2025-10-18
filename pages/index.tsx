import Head from "next/head";
import Link from "next/link";

export default function Home() {
  return (
    <>
      <Head>
        <title>Goobike Auto</title>
      </Head>
      <main style={{ padding: "2rem", fontFamily: "system-ui, sans-serif", maxWidth: "720px" }}>
        <h1>Goobike Auto 管理</h1>
        <p>
          このアプリケーションは Google Apps Script と連携するコードの管理リポジトリです。
          グーバイクへの登録作業を支援するスクリプトや説明ページをまとめています。
        </p>

        <section style={{ marginTop: "2rem" }}>
          <h2>ツール一覧</h2>
          <ul style={{ lineHeight: 1.8 }}>
            <li>
              <Link href="/autofill">スプレッドシート → GooBike フォーム自動入力スクリプト</Link>
            </li>
          </ul>
        </section>
      </main>
    </>
  );
}
