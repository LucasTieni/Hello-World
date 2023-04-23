import { Container } from 'components/container'
import { useRequest } from 'hooks/requests'
import Link from 'next/link'
import * as S from './styles'

export default function Footer() {
  const { data: footer, error } = useRequest(`/footers`)
  return (
    <S.Footer>
      <S.Widgets>
        <Container>
          <ul>
            {footer?.data?.map((item, i) => (
              <li key={item.attributes.titulo + i}>
                <span>{item.attributes.titulo}</span>
                <Link href={item.attributes.url} target="_blank">
                  {item.attributes.link}
                </Link>
              </li>
            ))}
          </ul>
        </Container>
      </S.Widgets>

      <S.Copyright>
        <p>
          &copy; 2023 F.A. Colchões - Todos os direitos reservados -
          <Link href="https://www.facolchoes.com.br" target="_blank">
            www.facolchoes.com.br
          </Link>
        </p>
      </S.Copyright>
    </S.Footer>
  )
}
