import Link from 'next/link'
import * as S from './styles'

export default function Menu() {
  return (
    <S.Menu>
      <ul>
        <li>
          <Link href="/">Início</Link>
        </li>
        <li>
          <Link href="/sobre-a-fa">Sobre</Link>
        </li>
        <li>
          <Link href="/campanhas">Campanhas</Link>
        </li>
        <li>
          <Link href="/produtos">Produtos</Link>
        </li>
        <li>
          <Link href="/treinamentos">Treinamentos</Link>
        </li>
        <li>
          <Link href="/eventos">Eventos</Link>
        </li>
        <li>
          <Link href="/noticias">Notícias</Link>
        </li>
      </ul>
    </S.Menu>
  )
}
