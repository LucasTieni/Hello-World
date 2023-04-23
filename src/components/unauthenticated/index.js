import Link from 'next/link'
import * as S from './styles'

export default function Unauthenticated() {
  return (
    <S.Unauthenticated>
      <p>Usuário não autenticado. Clique aqui para realizar o login: </p>
      <Link href="/login">
        <S.Button>Login</S.Button>
      </Link>
    </S.Unauthenticated>
  )
}
