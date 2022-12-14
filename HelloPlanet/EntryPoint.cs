// Copyright (C) 2013 InfoVista S.A. All rights reserved.

namespace HelloPlanet
{
    public class EntryPoint
    {
        /// <summary>
        /// Provides an entry point into HelloPlanet
        /// </summary>
        public void EnterSesame()
        {
            using (HelloPlanetForm form = new HelloPlanetForm())
            {

                form.ShowDialog();
            }
        }
    }
}
