import React, { useState } from 'react'
import Link from 'next/link'
import * as S from './styles'

export default function Nav() {
  // const { data: session, status } = useSession()
  const [open, setOpen] = useState({ transform: 'translate(-100%)' })
  return (
    <>
      <S.Menu css={open}>
        <S.Close onClick={() => setOpen({ transform: 'translate(-100%)' })}>
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24">
            <path fill="none" d="M0 0h24v24H0z" />
            <path d="M12 10.586l4.95-4.95 1.414 1.414-4.95 4.95 4.95 4.95-1.414 1.414-4.95-4.95-4.95 4.95-1.414-1.414 4.95-4.95-4.95-4.95L7.05 5.636z" />
          </svg>
        </S.Close>
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
      <S.Nav>
        <S.Mobile onClick={() => setOpen({ transform: 'translate(0)' })}>
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24">
            <path fill="none" d="M0 0h24v24H0z" />
            <path d="M3 4h18v2H3V4zm0 7h18v2H3v-2zm0 7h18v2H3v-2z" />
          </svg>
        </S.Mobile>
        <S.Brand>
          <Link href="/">
            <img src="../fa-logo.webp" alt="F.A. Colchões" />
          </Link>
        </S.Brand>
        <Link href={'/perfil'}>
          {/* <S.User>
            {status === 'authenticated' && (
              <>
                Olá,
                <span>
                  {'\u00A0'}
                  {session?.user?.username}!
                </span>
              </>
            )}
            {status !== 'authenticated' && <S.Button>Login</S.Button>}
          </S.User> */}
          <S.Avatar>
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24">
              <path fill="none" d="M0 0h24v24H0z" />
              <path d="M4 22a8 8 0 1 1 16 0h-2a6 6 0 1 0-12 0H4zm8-9c-3.315 0-6-2.685-6-6s2.685-6 6-6 6 2.685 6 6-2.685 6-6 6zm0-2c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4z" />
            </svg>
          </S.Avatar>
        </Link>
      </S.Nav>
    </>
  )
}
